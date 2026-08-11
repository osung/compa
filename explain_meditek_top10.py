# -*- coding: utf-8 -*-
"""MEDITEK TOP10 매칭 결과 → 보고서 입력 JSON(상세 추천 근거 + 표용 매칭 근거 생성).

match_meditek_top10.py 는 순위 선정까지만 하므로(적합도·우수성), 보고서에 필요한 서술은
여기서 35B 로 생성한다. 텍스트는 모두 모델이 쓰고, 이 스크립트는 프롬프트 구성과
체크포인트 관리만 한다.

  · 추천근거_상세 : 4섹션. 기업이 '보유한' 기술이 있으면 사양 적합성 대신
                   '기보유기술 보강 가능성' 섹션을 쓴다(MEDITEK 원본에 수요기술 사양이 없음).
  · 판단근거     : 표에 싣는 긍정형 한 문장(적합도 채점 단계의 비판적 reason 을 쓰지 않는다).

수요기술은 기업이 확보하려는(미보유) 기술, 기보유기술은 이미 보유한 역량이므로
payload 에서 각각 company.수요기술_* / company.description 으로 분리해 넣는다
(둘을 섞으면 '이미 보유했다'는 서술 오류가 난다).

사용: python explain_meditek_top10.py [--tag MEDITEK] [--only 3,7] [--limit 0]
산출: MEDITEK_TOP10_보고서.json (+ top10_<tag>_{explain,reason}_ckpt.json)
"""
import argparse
import json
import os
import re
import time

import pandas as pd

import compa_match as cm
import match_meditek_top10 as mt

TOP10_PKL = "MEDITEK_TOP10.pkl"
OUT_JSON = "MEDITEK_TOP10_보고서.json"

# ---- 상세 근거 섹션 구성 ----------------------------------------------------
EX_DEM = ["연관성", "수요 충족 가능성", "추천 과제의 우수성", "유사 사례 및 실적"]
EX_HOLD = ["연관성", "기보유기술 보강 가능성", "추천 과제의 우수성", "유사 사례 및 실적"]

_GUIDE = {
    "연관성": ("company.description(이 기업이 '이미 보유한' 기술·사업내용)과 company.수요기술명·"
            "수요기술_내용(이 기업이 '확보하려는', 아직 보유하지 않은 기술)을 구분해 읽고, "
            "이 기업의 기술적 위치에 비추어 과제의 목표·내용이 어떤 지점에서 맞닿는지 설명. "
            "수요기술을 기업이 이미 보유한 역량으로 서술하지 말 것. 반대로 description 에 있는 "
            "보유기술은 기업이 실제로 가진 것으로 서술해도 된다. 둘 중 한쪽이 비어 있으면 "
            "있는 쪽만 근거로 쓰고 없는 역량을 임의로 가정하지 말 것"),
    "수요 충족 가능성": ("과제가 company.수요기술_내용의 요구를 어떤 방식으로 충족·해결할 수 있는지, "
                  "과제의 기술 요소와 수요 항목을 대응시켜 구체적으로 설명. 완전히 일치하지 "
                  "않는 항목은 어느 범위까지 기여할 수 있는지로 서술"),
    "기보유기술 보강 가능성": ("과제의 기술이 company.description 의 보유기술을 어떻게 보강·고도화하거나 "
                     "새로운 응용으로 확장할 수 있는지 설명. 보유기술의 어느 구성요소(소재·공정·"
                     "알고리즘·계측 등)에 결합되는지, 그 결합이 성능·적용범위·신뢰성 측면에서 "
                     "무엇을 개선하는지 인과적으로 서술"),
    "추천 과제의 우수성": (cm._EX_GUIDE["추천 과제의 우수성"]
                   + ". 특허·논문 건수와 상위비율이 주어지면 그 수치가 뜻하는 강점을 함께 서술"),
    "유사 사례 및 실적": cm._EX_GUIDE["유사 사례 및 실적"],
}

# 표용 한 문장 근거: 보유기술 보강 접점도 허용(원 프롬프트는 수요 충족 관점만 전제)
_MR_EXTRA = ("\n6) 기업이 '이미 보유한 기술'이 함께 주어지면, 과제가 그 보유기술을 보강·고도화·"
             "확장하는 접점도 매칭 근거로 쓸 수 있다(예: '~ 기술이 보유 플랫폼 고도화에 활용 가능').")


def log(*a):
    print(*a, flush=True)


# ---- 표기 정리 --------------------------------------------------------------
# 원본 hwp 에서 넘어온 '3 D 배양'·'( PDC)' 같은 잔재가 모델 출력에도 옮겨 붙는다.
# 숫자-단위 사이 공백만 붙이되, 조사/구두점이 뒤따르는 경우로 한정해 '3 개발' → '3개발'
# 같은 오수정을 막는다.
_JOSA = r"(?=[의을를이가와과로에도만은는,·\s.)\]]|$)"
_NUM_SAFE = re.compile(r"(\d)\s+(차원|개월|시간|%|D(?![A-Za-z가-힣])|nm|㎛|mm|cm)")
_NUM_JOSA = re.compile(r"(\d)\s+(건|년|명|개|종|배|차)" + _JOSA)
_PAREN_L = re.compile(r"\(\s+")
_PAREN_R = re.compile(r"\s+\)")


def polish(s):
    """숫자-단위·괄호 공백 잔재 정리 + 마침표 뒤 공백 보정."""
    s = str(s or "")
    s = _PAREN_L.sub("(", _PAREN_R.sub(")", s))
    s = _NUM_SAFE.sub(r"\1\2", s)
    s = _NUM_JOSA.sub(r"\1\2", s)
    return cm.normalize_spacing(s)


def hold_body(unit):
    """기보유기술 합본에서 [기술명] 섹션을 뺀 본문(기술명은 별도 필드로 넣는다)."""
    return mt._HOLD_NAME_SEC.sub("", unit["기보유기술 내용"]).strip()


def payload_for(unit, kws, proj):
    demand = {
        "기업명": unit["기업명"],
        "수요기술명": unit["수요기술명"],
        "수요기술 내용": unit["수요기술 내용"],
        "수요기술 사양": "",
        "예상 적용 제품 및 서비스": "",
        "keywords": kws,
    }
    body = hold_body(unit)
    if body:                                   # 보유기술 = 기업의 실제 역량 → description
        demand["기업설명문"] = body[:2200]
        demand["desc_ok"] = 1
        demand["desc_issue"] = ""
    p = cm.build_demand_payload(demand, proj)
    fmt = EX_HOLD if body else EX_DEM
    if body:
        p["company"]["보유기술명"] = unit["기보유기술명"]
        p["company"]["보유기술_구분"] = " / ".join(
            x for x in (unit["기술유형"], unit["기술분야"]) if x)
    p["output_requirements"].update({
        "format": fmt,
        "section_guide": {k: _GUIDE[k] for k in fmt},
        "must_cover_수요기술_사양": False,
    })
    return p, fmt


def gen_detail(unit, kws, proj):
    """4섹션 상세 근거 → 한 셀 텍스트."""
    p, fmt = payload_for(unit, kws, proj)
    out = cm.stream_explanation(cm.build_messages(p, direction="company"),
                                max_tokens=1400, temperature=0.2, top_p=0.9,
                                expected_keys=fmt)
    secs = cm.parse_sections(out, tuple(fmt))
    parts = [f"[{k}] {secs.get(k, '').strip()}" for k in fmt if secs.get(k, "").strip()]
    return cm.normalize_spacing("\n\n".join(parts) if parts else out.strip())


def gen_reason(unit, proj):
    """표 '매칭 근거' 한 문장(긍정형). 실패 시 빈 문자열."""
    blocks = []
    if unit["수요기술 내용"]:
        blocks.append(f"수요기술명: {unit['수요기술명']}\n"
                      f"수요기술 내용: {unit['수요기술 내용'][:600]}")
    body = hold_body(unit)
    if body:
        blocks.append(f"이미 보유한 기술: {unit['기보유기술명']}\n"
                      f"보유기술 내용: {body[:600]}")
    pj = f"과제명: {proj.get('과제명','')}\n과제설명: {str(proj.get('설명','') or '')[:600]}"
    user = ("[기업 기술수요]\n" + "\n\n".join(blocks)
            + f"\n\n[R&D 과제]\n{pj}\n\n매칭 근거(한 문장):")
    msgs = ([{"role": "system", "content": cm._MR_SYS + _MR_EXTRA}]
            + cm._MR_FEWSHOT + [{"role": "user", "content": user}])
    try:
        out = cm.stream_explanation(msgs, max_tokens=120, temperature=0.0, top_p=1.0).strip()
    except Exception as e:
        log(f"      ! 매칭근거 실패: {e}")
        return ""
    out = out.strip().strip('"').strip("'").split("\n")[0].strip()
    if out.startswith("매칭 근거"):
        out = out.split(":", 1)[-1].strip()
    return cm.normalize_spacing(out[:120])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=TOP10_PKL)
    ap.add_argument("--tag", default="MEDITEK")
    ap.add_argument("--out", default=OUT_JSON)
    ap.add_argument("--only", default="", help="처리할 번호/관리번호(쉼표 구분)")
    ap.add_argument("--limit", type=int, default=0, help="처리 기업 수 제한(0=전체)")
    ap.add_argument("--no-llm", action="store_true",
                    help="LLM 호출 없이 체크포인트에 있는 근거만으로 JSON 재조립")
    a = ap.parse_args()

    df = pd.read_pickle(a.src)
    df["과제고유번호"] = df["과제고유번호"].astype(str)
    units, _ = mt.load_units()
    units.sort(key=mt.sort_key)
    ubykey = {u["키"]: u for u in units}
    kw_ckpt = json.load(open(os.path.join(cm.OUT_DIR, f"top10_{a.tag}_keywords_ckpt.json"),
                             encoding="utf-8"))

    ex_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_explain_ckpt.json")
    rs_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_reason_ckpt.json")
    ex = json.load(open(ex_path, encoding="utf-8")) if os.path.exists(ex_path) else {}
    rs = json.load(open(rs_path, encoding="utf-8")) if os.path.exists(rs_path) else {}
    log(f"체크포인트: 상세근거 {len(ex)}건 · 매칭근거 {len(rs)}건")

    # 과제 메타(논문·특허 실적) — 상세근거 프롬프트의 우수성/실적 근거
    import pickle
    with open(cm.PROJECT_META, "rb") as f:
        pmeta = pickle.load(f)

    todo = [u for u in units if u["키"] in set(df["기업명"].map(cm.norm_name))]
    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        todo = [u for u in todo if u["번호"] in keep or u["관리번호"] in keep]
    if a.limit:
        todo = todo[:a.limit]

    need = sum(1 for u in todo
               for _ in df[df["기업명"] == u["기업명"]].itertuples())
    log(f"대상 기업 {len(todo)}건 · 추천 {need}건")
    if not a.no_llm:
        log("· 35B 로드…")
        cm.load_model_blocking(progress_cb=lambda m: log("  " + m))

    best, t0, done = {}, time.time(), 0
    for n, u in enumerate(todo, 1):
        key = u["번호"] or u["관리번호"]
        g = df[df["기업명"] == u["기업명"]].sort_values("rank")
        kws = kw_ckpt.get(u["키"], [])
        tops = []
        for r in g.to_dict("records"):
            pid = str(r["과제고유번호"])
            ck = f"{key}::{pid}"
            meta = pmeta.get(pid, {}) if isinstance(pmeta, dict) else {}
            proj = {
                "pid": pid, "과제명": r["과제명"], "설명": r["과제설명문"],
                "유망성": round(float(r["유망성점수"]), 1), "수행기관": r["과제수행기관"],
                "키워드": [], "논문명": meta.get("논문명_리스트") or [],
                "특허명": meta.get("특허명_리스트") or [],
                "논문건수": int(r["논문건수"]), "특허건수": int(r["특허건수"]),
                "총연구비_상위비율": meta.get("총연구비_상위비율"),
                "논문건수_상위비율": meta.get("논문건수_상위비율"),
                "특허건수_상위비율": meta.get("특허건수_상위비율"),
            }
            if not a.no_llm and ck not in ex:
                ex[ck] = gen_detail(u, kws, proj)[:cm.DESC_OUT]
                json.dump(ex, open(ex_path, "w", encoding="utf-8"), ensure_ascii=False)
            if not a.no_llm and ck not in rs:
                rs[ck] = gen_reason(u, proj)
                json.dump(rs, open(rs_path, "w", encoding="utf-8"), ensure_ascii=False)
            done += 1
            tops.append({
                "rank": int(r["rank"]), "과제고유번호": pid, "과제명": r["과제명"],
                "수행기관": r["과제수행기관"],
                "적합도": int(r["적합도"]), "수요충족": int(r["수요충족"]),
                "보유보강": int(r["보유보강"]), "최종점수": float(r["최종점수"]),
                "특허건수": int(r["특허건수"]), "논문건수": int(r["논문건수"]),
                "유망성점수": float(r["유망성점수"]), "순위구분": r["순위구분"],
                # 표에는 긍정형 한 문장, 없으면 채점 단계 근거로 폴백
                "판단근거": polish(rs.get(ck) or r["LLM근거"]),
                "과제설명문": r["과제설명문"],
                "추천근거_상세": polish(ex.get(ck, "")),
            })
        best[key] = {
            "번호": u["번호"], "관리번호": u["관리번호"], "기업명": u["기업명"],
            "소스": u["소스"], "키워드": cm.SEP.join(kws),
            "수요기술명": polish(u["수요기술명"]), "수요기술 내용": polish(u["수요기술 내용"]),
            "기보유기술명": polish(u["기보유기술명"]), "기보유기술 내용": polish(hold_body(u)),
            "기술유형": u["기술유형"], "기술분야": u["기술분야"],
            "top10": tops,
        }
        el = time.time() - t0
        log(f"      [{key}] {u['기업명']}({u['소스']}): {len(tops)}건 근거 완료 "
            f"[{n}/{len(todo)}] {el/60:.1f}분 경과"
            + (f" · 남은 예상 {el/done*(need-done)/60:.0f}분" if done and not a.no_llm else ""))

    if os.path.exists(a.out):
        os.replace(a.out, a.out + ".bak")
    json.dump(best, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    n_rec = sum(len(v["top10"]) for v in best.values())
    n_ex = sum(1 for v in best.values() for t in v["top10"] if t["추천근거_상세"])
    log(f"\n✔ {a.out}: 기업 {len(best)}건 · 추천 {n_rec}건 · 상세근거 {n_ex}건 "
        f"· 중복제외 과제 {len({t['과제고유번호'] for v in best.values() for t in v['top10']})}개")
    if n_ex < n_rec:
        log(f"! 상세근거 누락 {n_rec - n_ex}건 — 다시 실행하면 누락분만 생성한다")


if __name__ == "__main__":
    main()
