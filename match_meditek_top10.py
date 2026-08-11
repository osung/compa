# -*- coding: utf-8 -*-
"""수요기술 + 기보유기술 통합 매칭 → 국가 R&D 과제 TOP10.

기존 match_meditek.py 는 '확보하려는 수요기술'만 질의로 썼다. 이 스크립트는 기업이
이미 보유한 기술·사업내용(MEDITEK_기보유기술.pkl)까지 함께 매칭에 활용한다.

  ① 키워드   : 수요기술·기보유기술을 가리지 않고 한 문서로 합쳐 핵심 키워드 추출(35B)
  ② 후보검색 : 키워드 문자열 → pro-sroberta → 과제 임베딩 코사인 상위 N건
  ③ LLM평가  : 후보별로 (가) 수요기술을 충족·해결하는가 (나) 기보유기술을 보강·고도화하는가
               를 각각 0~100 으로 채점 → 적합도 = 둘 중 높은 값(둘 중 하나만 통해도 매칭)
  ④ TOP10    : 적합도 + 과제 우수성(특허건수 ≫ 논문건수·유망성)의 가중합으로 순위 결정

과제 우수성에서 '보유 특허 수'를 가장 큰 가중치로 둔다(--w-pat). 다만 우수성만으로
무관한 과제가 올라오지 않도록 적합도 하한(--min-fit) 통과분을 먼저 채우고, 통과분이
TOP10 에 미달할 때만 나머지를 '보충'(순위구분 컬럼)으로 뒤에 붙인다. 보충분은 최종점수가
더 높아도 통과분보다 앞서지 않는다.

코퍼스 필터는 기존 최종본과 동일(연구수행주체 4종 ∧ 제출년도 2020~).
LLM 은 compa_match 의 백엔드를 그대로 사용(MLX: mlx-community/Qwen3.5-35B-A3B-4bit).

사용: python match_meditek_top10.py [--tag MEDITEK] [--n-cand 200] [--final 10]
산출: <tag>_TOP10.xlsx / .pkl  (체크포인트: top10_<tag>_{keywords,scores}_ckpt.json)
"""
import argparse
import json
import os
import pickle
import re
import time

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer

import compa_match as cm
from rematch_filtered import ALLOW, EMB_FILE, YEAR_MIN

DEMANDS_PKL = "MEDITEK_수요기술.pkl"
FIRMS_PKL = "MEDITEK_기보유기술.pkl"

N_CAND = 200               # pro-sroberta 코사인 상위 후보 수(LLM 평가 대상)
FINAL = 10                 # 최종 추천 수
MIN_FIT = 40               # 적합도 하한(미달은 우수성이 높아도 '보충'으로만 들어감)
# 최종 점수 가중치(합 1.0). 우수성 중에서는 특허를 가장 크게 둔다.
W_FIT, W_PAT, W_PAP, W_PROM = 0.60, 0.24, 0.06, 0.10
# 건수 정규화 기준(0=코퍼스 분위수 자동). 특허는 변별력 확보를 위해 P99, 논문은 P95.
PAT_REF, PAP_REF = 0.0, 0.0
DEM_MAX, HOLD_MAX = 1200, 1600     # 프롬프트에 넣는 수요/기보유 본문 절단 길이
RETRY = 2                          # 배치 응답에서 빠진 후보 재질의 횟수


def log(*a):
    print(*a, flush=True)


def _to_int(v):
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
        return 0


# ---------------------------------------------------------------- 매칭 단위
def load_units(dsrc=DEMANDS_PKL, fsrc=FIRMS_PKL):
    """수요기술 DF + 기보유기술 DF → 기업 단위 매칭 유닛 리스트.

    기업명(법인격 표기 무시)으로 두 표를 합친다. 수요기술만 있는 기업, 기보유기술만
    있는 기업, 둘 다 있는 기업을 모두 싣고 둘 다 비면 제외한다.
    """
    dm = pd.read_pickle(dsrc)
    dm["번호"] = dm["번호"].astype(str)
    fm = (pd.read_pickle(fsrc) if os.path.exists(fsrc)
          else pd.DataFrame(columns=["기업명_norm"]))
    firms = {r["기업명_norm"]: r for r in fm.to_dict("records")}

    units, used = [], set()
    for r in dm.to_dict("records"):
        key = r["기업명_norm"]
        f = firms.get(key, {})
        used.add(key)
        units.append(_unit(key, r, f))
    for key, f in firms.items():                  # 수요 행이 없는 기업(기보유기술만)
        if key not in used:
            units.append(_unit(key, {}, f))

    keep = [u for u in units if u["소스"] != "없음"]
    skip = [u for u in units if u["소스"] == "없음"]
    return keep, skip


def _unit(key, d, f):
    dem = str(d.get("수요기술 내용", "") or "")
    hold = str(f.get("기보유기술 내용", "") or "")
    src = ("수요+기보유" if dem and hold else "수요" if dem else "기보유" if hold else "없음")
    return {
        "키": key,
        "기업명": str(d.get("기업명") or f.get("기업명") or ""),
        "번호": str(d.get("번호", "") or ""),
        "관리번호": str(f.get("관리번호", "") or ""),
        "수요기술명": str(d.get("수요기술명", "") or ""),
        "수요기술 내용": dem,
        "기보유기술명": str(f.get("기술명", "") or ""),
        "기보유기술 내용": hold,
        "기술유형": str(f.get("기술유형", "") or ""),
        "기술분야": str(f.get("기술분야", "") or ""),
        "소스": src,
    }


def unified_doc(u):
    """키워드 추출용 통합 문서 — 수요기술과 기보유기술을 구분 없이 한 문서로."""
    parts = []
    if u["수요기술 내용"]:
        parts.append(f"[확보하려는 수요기술] {u['수요기술명']}\n{u['수요기술 내용']}")
    if u["기보유기술 내용"]:
        parts.append(f"[이미 보유한 기술·사업내용] {u['기보유기술명']}\n{u['기보유기술 내용']}")
    return "\n\n".join(parts)


# 통합 문서의 라벨·연결어 조각(수요 본문이 짧아 LLM 키워드가 KW_MIN 미만이면
# extract_keywords_doc 의 규칙기반 폴백이 라벨까지 명사로 긁어온다).
_KW_DROP = {"확보하려는", "수요기술", "이미", "보유한", "사업내용", "기보유기술",
            "기술·사업내용", "관심이", "있으며", "해당"}


def filter_keywords(kws):
    """라벨 조각 제거 + 다른 키워드에 통째로 포함된 파편 제거.

    폴백 명사추출은 '적층 성형' 과 함께 '적층'·'디지털' 같은 조각도 넣는데, 이런
    부분문자열은 질의에 정보를 더하지 않고 노이즈만 된다. 단어 경계로 포함될 때만
    떨어내 'RI' 가 'PRINT' 에 걸리는 식의 오제거를 막는다.
    """
    kws = [k for k in (str(x).strip() for x in kws) if k and k not in _KW_DROP]
    out = []
    for k in kws:
        pat = re.compile(rf"(?<!\w){re.escape(k)}(?!\w)")
        if any(o != k and pat.search(o) for o in kws):
            continue
        out.append(k)
    return out


def keywords_of(u):
    """통합 문서(수요기술+기보유기술) → 핵심 키워드."""
    return filter_keywords(cm.extract_keywords_doc(unified_doc(u)))


def sort_key(u):
    """수요 번호 순 → 그다음 기보유기술만 있는 기업(관리번호 순)."""
    if u["번호"]:
        return (0, float(u["번호"]))
    m = re.search(r"(\d+)$", u["관리번호"])
    return (1, float(m.group(1)) if m else 0.0)


# ---------------------------------------------------------------- 코퍼스
def build_corpus():
    """필터 코퍼스(임베딩 + 과제 메타). SBERT 검색만 쓰므로 BM25·기업설명문은 싣지 않는다."""
    t0 = time.time()
    log(f"· 임베딩 로드+필터 (주체 {sorted(ALLOW)} ∧ 제출년도>={YEAR_MIN})…")
    pdf = pd.read_pickle(EMB_FILE)
    y = pd.to_numeric(pdf["제출년도"], errors="coerce")
    pdf = pdf[pdf["연구수행주체"].isin(ALLOW) & (y >= YEAR_MIN)].reset_index(drop=True)
    log(f"  필터 통과 {len(pdf)}건 ({time.time()-t0:.0f}s)")

    c = {
        "pid": pdf["과제고유번호"].astype(str).values,
        "pname": pdf["과제명"].astype(str).values,
        "pdesc": pdf["과제설명문"].astype(str).values,
        "pkw": pdf["키워드_리스트"].values,
        "promise": pdf["유망성점수"].values.astype(np.float32),
    }
    M = np.vstack([np.asarray(e, dtype=np.float32) for e in pdf["norm_embed"].values])
    M /= np.linalg.norm(M, axis=1, keepdims=True)
    c["M"] = M
    del pdf

    log("· 과제 메타(수행기관·특허·논문) 로드…")
    with open(cm.PROJECT_META, "rb") as fp:
        pmeta = pickle.load(fp)
    c["pmeta"] = pmeta
    c["org"] = np.array([str(pmeta.get(p, {}).get("과제수행기관명", "")) for p in c["pid"]])
    c["pat"] = np.array([_to_int(pmeta.get(p, {}).get("특허건수")) for p in c["pid"]],
                        dtype=np.int32)
    c["pap"] = np.array([_to_int(pmeta.get(p, {}).get("논문건수")) for p in c["pid"]],
                        dtype=np.int32)
    log(f"  코퍼스 준비 완료 ({time.time()-t0:.0f}s)")
    return c


def excellence_refs(c, pat_ref=0.0, pap_ref=0.0):
    """특허·논문건수 정규화 기준. 분포를 함께 출력해 근거를 남긴다.

    코퍼스의 85%가 특허 0건이라 P95(=3건)를 기준으로 삼으면 3건에서 만점에 포화되어
    변별력이 사라진다. 특허는 '매우 중요한 점수'이므로 기준을 P99 로 올려 다건 보유
    과제가 실제로 더 높은 점수를 받게 한다. 논문은 보조 항목이라 P95 를 쓴다.
    """
    refs, given = {}, {"pat": pat_ref, "pap": pap_ref}
    for k, lab, pct in (("pat", "특허건수", 99), ("pap", "논문건수", 95)):
        v = c[k]
        q = dict(zip((50, 90, 95, 99), np.percentile(v, [50, 90, 95, 99])))
        refs[k] = max(float(given[k] or q[pct]), 1.0)
        log(f"  {lab} 분포: 0건 {int((v == 0).sum())}건({(v == 0).mean()*100:.0f}%) · "
            f"P50={q[50]:.0f} P90={q[90]:.0f} P95={q[95]:.0f} P99={q[99]:.0f} max={v.max()} "
            f"→ 정규화 기준 {refs[k]:.0f}"
            + ("(지정)" if given[k] else f"(P{pct})"))
    return refs


def norm_count(x, ref):
    """건수 → 0~1. 상위 소수 과제가 점수를 독식하지 않도록 로그 스케일 후 P95 로 정규화."""
    return float(min(1.0, np.log1p(max(0, x)) / np.log1p(ref)))


# ---------------------------------------------------------------- LLM 평가
_SYS = ("당신은 기업의 기술과 국가 R&D 과제의 연계 가능성을 평가하는 기술이전 전문가입니다. "
        "각 과제가 (가) 기업이 확보하려는 수요기술을 충족·해결하는지, "
        "(나) 기업이 이미 보유한 기술을 보강·고도화하는지를 각각 냉정하게 평가합니다.")

_BANDS = ("  90-100: 동일 기술/제품을 직접 해결·보강하는 매우 높은 적합\n"
          "  70-89 : 핵심 기술·응용 분야가 상당 부분 부합\n"
          "  40-69 : 일부 요소만 관련(부분 적합)\n"
          "  10-39 : 분야만 유사하거나 약하게 관련\n"
          "  0-9   : 사실상 무관\n")


def _guide(has_dem, has_hold):
    """평가 관점 안내 — 기업에 실제로 있는 정보(수요/기보유)에 대해서만 채점을 요구한다."""
    keys, expl = [], []
    if has_dem:
        keys.append('"수요충족": <0-100 정수>')
        expl.append("  수요충족: 이 과제가 [확보하려는 수요기술]을 충족·해결하는 정도")
    if has_hold:
        keys.append('"보유보강": <0-100 정수>')
        expl.append("  보유보강: 이 과제가 [이미 보유한 기술]을 보강·고도화하거나 "
                    "새 응용으로 확장하는 데 쓰일 수 있는 정도")
    both = has_dem and has_hold
    head = ("각 과제를 아래 두 관점에서 각각 0~100 정수로 평가하세요.\n" if both
            else "각 과제를 아래 관점에서 0~100 정수로 평가하세요.\n")
    indep = ("두 관점의 점수는 서로 독립입니다(한쪽이 낮아도 다른 쪽이 높을 수 있습니다).\n"
             if both else "")
    return (head + "\n".join(expl) + "\n" + _BANDS
            + "키워드 표면 일치가 아니라 기술 내용·적용 제품 관점의 실질 적합도를 보세요.\n"
            + indep
            + "반드시 아래 JSON 형식으로만, 모든 과제에 대해 답하세요(reason 은 40자 이내 한국어):\n"
            '{"results": [{"id": <과제번호>, ' + ", ".join(keys)
            + ', "reason": "<근거>"}, ...]}')


# 기보유기술 합본 앞머리의 [기술명] 섹션(프롬프트에 '- 기술명:' 으로 이미 들어감)
_HOLD_NAME_SEC = re.compile(r"^\[기술명\]\n.*?(?=\n\n\[|\Z)", re.DOTALL)


def _unit_block(u):
    b = [f"[기업] {u['기업명']}"]
    if u["수요기술 내용"]:
        b.append("[확보하려는 수요기술]\n"
                 f"- 수요기술명: {u['수요기술명']}\n"
                 f"- 내용: {u['수요기술 내용'][:DEM_MAX]}")
    if u["기보유기술 내용"]:
        meta = " / ".join(x for x in (u["기술유형"], u["기술분야"]) if x)
        body = _HOLD_NAME_SEC.sub("", u["기보유기술 내용"]).strip()
        b.append("[이미 보유한 기술·사업내용]\n"
                 f"- 기술명: {u['기보유기술명']}"
                 + (f"\n- 구분: {meta}" if meta else "")
                 + f"\n- 내용: {body[:HOLD_MAX]}")
    return "\n\n".join(b)


def _cand_block(cands):
    lines = ["[평가 대상 R&D 과제 목록]"]
    for c in cands:
        kw = ", ".join(c["키워드"][:8])
        lines.append(f"{c['id']}. 과제명: {c['과제명']}\n"
                     f"   키워드: {kw}\n"
                     f"   설명: {c['과제설명문'][:cm.DESC_CAND]}")
    return "\n".join(lines)


def _parse(text, has_dem, has_hold):
    """{"results":[...]} → {id: (수요충족, 보유보강, reason)}. 없는 관점은 0."""
    m = re.search(r'\{.*"results".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("results", []) if isinstance(data, dict) else data
    except Exception:                              # 폴백: 객체 단위 정규식
        items = [json.loads(x) for x in re.findall(r'\{[^{}]*"id"[^{}]*\}', blob)
                 if _loadable(x)]
    out = {}
    for it in items:
        try:
            i = int(it["id"])
        except (KeyError, ValueError, TypeError):
            continue
        d = _clip(it.get("수요충족")) if has_dem else 0
        h = _clip(it.get("보유보강")) if has_hold else 0
        out[i] = (d, h, str(it.get("reason", "")).strip()[:120])
    return out


def _loadable(s):
    try:
        json.loads(s)
        return True
    except Exception:
        return False


def _clip(v):
    try:
        return max(0, min(100, int(round(float(v)))))
    except (TypeError, ValueError):
        return 0


def _score_batch(u, batch, guide, has_dem, has_hold):
    """후보 묶음 1회 채점 → {id: (수요충족, 보유보강, reason)}(응답 누락 id 는 빠진다)."""
    user = f"{_unit_block(u)}\n\n{_cand_block(batch)}\n\n{guide}"
    msgs = [{"role": "system", "content": _SYS}, {"role": "user", "content": user}]
    try:
        return _parse(cm.stream_explanation(msgs, max_tokens=cm.LLM_MAXTOK,
                                            temperature=0.0, top_p=1.0),
                      has_dem, has_hold)
    except Exception as e:
        log(f"      ! LLM 배치 실패: {e}")
        return {}


def score_candidates(u, cands, retry=RETRY):
    """후보 전체를 LLM_BATCH 단위로 이중 채점 → {id: (수요충족, 보유보강, reason)}.

    JSON 응답에서 일부 id 가 빠지는 일이 있어(긴 배치의 뒷부분), 누락분은 그 id 들만
    다시 물어 본다(retry 회). 끝까지 못 받은 후보는 0점 처리해 순위에서 사실상 제외된다.
    """
    has_dem, has_hold = bool(u["수요기술 내용"]), bool(u["기보유기술 내용"])
    guide = _guide(has_dem, has_hold)
    scores, miss = {}, 0
    for b in range(0, len(cands), cm.LLM_BATCH):
        batch = cands[b:b + cm.LLM_BATCH]
        parsed = _score_batch(u, batch, guide, has_dem, has_hold)
        for _ in range(retry):                    # 누락 id 만 재질의(id 번호는 유지)
            left = [c for c in batch if c["id"] not in parsed]
            if not left:
                break
            parsed.update(_score_batch(u, left, guide, has_dem, has_hold))
        for c in batch:
            if c["id"] in parsed:
                scores[c["id"]] = parsed[c["id"]]
            else:
                scores[c["id"]] = (0, 0, "(LLM 응답 누락)")
                miss += 1
    if miss:
        log(f"      ! LLM 응답 누락 {miss}/{len(cands)}건(재질의 {retry}회 후에도 없음 → 0점)")
    return scores


# ---------------------------------------------------------------- 산출
XLSX_WIDTH = {"번호": 6, "관리번호": 10, "기업명": 16, "소스": 11, "수요기술명": 30,
              "기보유기술명": 30, "키워드": 30, "rank": 5, "최종점수": 8, "적합도": 7,
              "수요충족": 8, "보유보강": 8, "특허건수": 8, "논문건수": 8, "유망성점수": 9,
              "특허점수": 8, "유사도_코사인": 12, "과제고유번호": 14, "과제명": 40,
              "과제수행기관": 22, "LLM근거": 40, "순위구분": 8, "과제설명문": 60}
WRAP_COLS = {"수요기술명", "기보유기술명", "키워드", "과제명", "LLM근거", "과제설명문"}


def save_xlsx(df, path):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    df.to_excel(path, index=False)
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    for j, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = XLSX_WIDTH.get(col, 14)
        if col in WRAP_COLS:
            for i in range(2, ws.max_row + 1):
                ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--demands", default=DEMANDS_PKL)
    ap.add_argument("--firms", default=FIRMS_PKL)
    ap.add_argument("--tag", default="MEDITEK")
    ap.add_argument("--only", default="", help="처리할 수요번호/관리번호(쉼표 구분)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--n-cand", type=int, default=N_CAND, help="SBERT 후보 수(LLM 평가 대상)")
    ap.add_argument("--final", type=int, default=FINAL)
    ap.add_argument("--min-fit", type=int, default=MIN_FIT)
    ap.add_argument("--w-fit", type=float, default=W_FIT)
    ap.add_argument("--w-pat", type=float, default=W_PAT)
    ap.add_argument("--w-pap", type=float, default=W_PAP)
    ap.add_argument("--w-prom", type=float, default=W_PROM)
    ap.add_argument("--pat-ref", type=float, default=PAT_REF,
                    help="특허건수 만점 기준(0=코퍼스 P99)")
    ap.add_argument("--pap-ref", type=float, default=PAP_REF,
                    help="논문건수 만점 기준(0=코퍼스 P95)")
    ap.add_argument("--dedupe-phase", action="store_true",
                    help="같은 사업의 연차 후속과제(…Ⅵ/Ⅶ, N차)를 한 건으로 보고 선정")
    a = ap.parse_args()

    units, skipped = load_units(a.demands, a.firms)
    units.sort(key=sort_key)
    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        units = [u for u in units if u["번호"] in keep or u["관리번호"] in keep]
    if a.limit:
        units = units[:a.limit]

    wsum = a.w_fit + a.w_pat + a.w_pap + a.w_prom
    log(f"대상 기업 {len(units)}건 "
        f"(수요+기보유 {sum(u['소스'] == '수요+기보유' for u in units)} · "
        f"수요만 {sum(u['소스'] == '수요' for u in units)} · "
        f"기보유만 {sum(u['소스'] == '기보유' for u in units)})")
    if skipped:
        log(f"제외(수요·기보유 모두 없음) {len(skipped)}건: "
            + ", ".join(f"{u['번호'] or u['관리번호']}:{u['기업명']}" for u in skipped))
    log(f"가중치 fit={a.w_fit} 특허={a.w_pat} 논문={a.w_pap} 유망성={a.w_prom} (합 {wsum:.2f}) "
        f"· 적합도 하한 {a.min_fit} · 후보 {a.n_cand} → TOP{a.final}")

    kw_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_keywords_ckpt.json")
    sc_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_scores_ckpt.json")
    kw_ckpt = json.load(open(kw_path, encoding="utf-8")) if os.path.exists(kw_path) else {}
    sc_ckpt = json.load(open(sc_path, encoding="utf-8")) if os.path.exists(sc_path) else {}
    if kw_ckpt or sc_ckpt:
        log(f"체크포인트 로드: 키워드 {len(kw_ckpt)}건 · 점수 {len(sc_ckpt)}건")

    log("· 35B 로드…")
    cm.load_model_blocking(progress_cb=lambda m: log("  " + m))

    log("· 키워드 추출(수요기술+기보유기술 통합)…")
    for n, u in enumerate(units, 1):
        if u["키"] not in kw_ckpt:
            kw_ckpt[u["키"]] = keywords_of(u)
            json.dump(kw_ckpt, open(kw_path, "w", encoding="utf-8"), ensure_ascii=False)
        u["키워드"] = kw_ckpt[u["키"]]
        log(f"      [{u['번호'] or u['관리번호']}] {u['기업명']}({u['소스']}): "
            f"키워드 {len(u['키워드'])}개 [{n}/{len(units)}]")

    log("· pro-sroberta 로드…")
    model = SentenceTransformer(cm.MODEL_DIR, device="cpu")
    model.eval()
    encode = cm.make_encoder(model)

    c = build_corpus()
    refs = excellence_refs(c, a.pat_ref, a.pap_ref)

    rows, t0 = [], time.time()
    for n, u in enumerate(units, 1):
        cos = np.clip(c["M"] @ encode(cm.SEP.join(u["키워드"])), -1, 1)
        cand_idx = np.argsort(-cos)[:a.n_cand]
        cands = [{"id": j + 1, "과제명": c["pname"][i], "과제설명문": c["pdesc"][i],
                  "키워드": list(c["pkw"][i]) if isinstance(c["pkw"][i], (list, tuple)) else []}
                 for j, i in enumerate(cand_idx)]

        # 체크포인트는 후보 순서에 흔들리지 않게 과제고유번호(pid) 키로 저장
        got = sc_ckpt.get(u["키"], {})
        need = [x for x in cands if str(c["pid"][cand_idx[x["id"] - 1]]) not in got]
        if need:
            log(f"      [{u['번호'] or u['관리번호']}] {u['기업명']}: "
                f"LLM 이중채점 {len(need)}/{len(cands)}건…")
            sc = score_candidates(u, need)
            for x in need:
                got[str(c["pid"][cand_idx[x["id"] - 1]])] = list(sc[x["id"]])
            sc_ckpt[u["키"]] = got
            json.dump(sc_ckpt, open(sc_path, "w", encoding="utf-8"), ensure_ascii=False)

        # 점수 합성
        scored = []
        for j, i in enumerate(cand_idx):
            d, h, reason = got.get(str(c["pid"][i]), (0, 0, ""))
            fit = max(d, h)
            pat_n = norm_count(int(c["pat"][i]), refs["pat"])
            pap_n = norm_count(int(c["pap"][i]), refs["pap"])
            prom_n = float(np.clip(c["promise"][i] / 100.0, 0, 1))
            total = 100.0 * (a.w_fit * fit / 100.0 + a.w_pat * pat_n
                             + a.w_pap * pap_n + a.w_prom * prom_n) / wsum
            scored.append({"i": int(i), "fit": fit, "d": d, "h": h, "reason": reason,
                           "pat_n": pat_n, "total": total, "cos": float(cos[i])})
        scored.sort(key=lambda x: (-x["total"], -x["fit"], -x["cos"]))

        # 적합도 하한 통과분에서 과제명 중복 제거 후 TOP-final
        # (하한 미달은 '보충'으로 뒤에만 붙는다 — 최종점수가 더 높아도 앞서지 않는다)
        seen, sel = set(), []
        for pool in (True, False):               # 1차: 하한 통과분 / 2차: 부족분 보충
            for x in scored:
                if (x["fit"] >= a.min_fit) != pool:
                    continue
                tk = cm.title_key(c["pname"][x["i"]], a.dedupe_phase)
                if tk in seen:
                    continue
                seen.add(tk)
                sel.append(dict(x, 순위구분="적합" if pool else "보충"))
                if len(sel) >= a.final:
                    break
            if len(sel) >= a.final:
                break
        n_pass = sum(1 for x in scored if x["fit"] >= a.min_fit)
        if n_pass < a.final:
            log(f"      ! 적합도 {a.min_fit} 이상 후보 {n_pass}건 < TOP{a.final} "
                f"→ 부족분은 적합도 순으로 보충(하한통과=False 표시)")

        for rank, x in enumerate(sel, 1):
            i = x["i"]
            rows.append({
                "번호": u["번호"], "관리번호": u["관리번호"], "기업명": u["기업명"],
                "소스": u["소스"], "수요기술명": u["수요기술명"],
                "기보유기술명": u["기보유기술명"], "키워드": cm.SEP.join(u["키워드"]),
                "rank": rank, "최종점수": round(x["total"], 2), "적합도": x["fit"],
                "수요충족": x["d"], "보유보강": x["h"],
                "특허건수": int(c["pat"][i]), "논문건수": int(c["pap"][i]),
                "유망성점수": round(float(c["promise"][i]), 2),
                "특허점수": round(x["pat_n"], 3),
                "유사도_코사인": round(x["cos"], 6),
                "과제고유번호": str(c["pid"][i]), "과제명": c["pname"][i],
                "과제수행기관": c["org"][i], "LLM근거": x["reason"],
                "순위구분": x["순위구분"],
                "과제설명문": c["pdesc"][i][:cm.DESC_OUT],
            })
        best = sel[0] if sel else None
        log(f"      [{u['번호'] or u['관리번호']}] {u['기업명']}: 후보 {len(cands)} "
            f"→ 하한통과 {n_pass} → TOP{len(sel)}"
            + (f" (1위 {best['total']:.1f}점 적합도 {best['fit']} "
               f"특허 {int(c['pat'][best['i']])}건)" if best else "")
            + f" [{n}/{len(units)}]")

    df = pd.DataFrame(rows)
    xlsx = os.path.join(cm.OUT_DIR, f"{a.tag}_TOP10.xlsx")
    df.to_pickle(os.path.join(cm.OUT_DIR, f"{a.tag}_TOP10.pkl"))
    save_xlsx(df, xlsx)
    log(f"\n✔ 완료 ({len(units)}기업 · {len(df)}행 · {time.time()-t0:.0f}s) → {xlsx} (+ .pkl)")


if __name__ == "__main__":
    main()
