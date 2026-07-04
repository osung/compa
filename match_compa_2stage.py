# -*- coding: utf-8 -*-
"""COMPA 진성수요 시트의 기업 수요에 대해 단일 SBERT 매칭 + LLM 적합도 평가로
국가 R&D 과제를 추천한다.

배경
----
구글시트(COMPA_진성수요)의 각 기업 수요 1건(=시트 한 행, 담당자별)에 대해 아래
파이프라인을 수행한다.

파이프라인 (기업 수요 1건)
  [키워드 추출] 수요기술명 + 수요기술 내용 + 수요기술 사양 → LLM(mve)으로 핵심 기술
      키워드 10~20개 추출 → ';' 로 join → '키워드' 컬럼. (KW_CKPT 체크포인트)
  [임베딩]      pro-sroberta 로 '키워드' 문자열 인코딩 → L2 정규화 → query 벡터 q.
  [SBERT 매칭]  q 와 과제 임베딩(public_RnD_embeddings_pro_260601_with_desc.pkl 의
      norm_embed)의 코사인 유사도 → 상위 TOPK(기본 100) 과제(= LLM 평가 대상).
  [LLM 적합도]  TOPK 후보 각각이 수요기술명·수요기술 내용·수요기술 사양을 실제로
      충족하는지 LLM 으로 0~100 정량 점수화(배치) → 점수 내림차순 상위 FINAL(기본 5)
      선정. (LLM_CKPT)
  [상세 근거]   최종 FINAL 개에 한해서만, matching_viewer_explain 기반 4섹션 상세
      추천 근거(수요기술 사양 적합성 포함)를 생성. (EXPLAIN_CKPT)

출력 (원본 시트는 수정하지 않음)
  - OUT_KW_XLSX  : 대상 수요 전체 + '키워드' 컬럼 (키워드 추출 결과)
  - OUT_FINAL_XLSX : 기업별 최종 추천 FINAL 개 (LLM점수·LLM판단근거·추천근거_상세 포함)

LLM 점수/설명 함수는 match_compa_list.py 의 검증된 구현을 재사용한다.
"""
import os
import re
import json
import pickle
import argparse

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from sentence_transformers import SentenceTransformer

import user_input_keywords as uik          # (참고용; 본 스크립트는 LLM 키워드 추출 사용)
import matching_viewer_explain as mve
import match_compa_list as mcl              # llm_score / gen_explanation / build_demand_payload 재사용

# ---- 설정 ----------------------------------------------------------------
XLSX = "COMPA_진성수요_원본.xlsx"     # 입력(구글시트 다운로드본, 수정하지 않음)
ASSIGNEE_COL_IDX = 1                  # 담당자 컬럼(2번째, 헤더 'Unnamed: 1')
DEFAULT_ASSIGNEE = "이중연"           # 기본 대상 담당자
HEADER_ROW = 2                        # pd.read_excel header 행(0-base): 상단 제목/병합행 2줄 skip


# 산출물/체크포인트는 담당자별로 분리한다(번호가 담당자별로 1부터 시작해도 충돌 없음).
def kw_xlsx_path(a):
    return f"COMPA_{a}_키워드.xlsx"            # 1차 산출물: 키워드 추출 결과


def final_xlsx_path(a):
    return f"COMPA_{a}_최종추천.xlsx"          # 2차 산출물: 최종 추천 Top10


def kw_ckpt_path(a):
    return f"compa2stage_{a}_keywords_ckpt.json"     # {번호: [키워드,...]}


def llm_ckpt_path(a):
    # 후보 위치(id 1..topk) 기반 점수 → 단일 SBERT top100 후보에 맞춘 별도 체크포인트
    # (옛 2단계 후보로 매긴 compa2stage_{a}_llm_scores_ckpt.json 과 혼용 방지)
    return f"compa2stage_{a}_llm_scores_sbert100_ckpt.json"   # {번호: {id: [score,reason]}}


def ex_ckpt_path(a):
    return f"compa2stage_{a}_explain_ckpt.json"      # {번호::과제번호: text}

PROJECT_EMB = "public_RnD_embeddings_pro_260601_with_desc.pkl"   # 과제 임베딩(matching_viewer 과제 데이터)
COMPANY_EMB = "company_embeddings_pro_260514_with_desc.pkl"      # 기업 임베딩(matching_viewer 기업 데이터)
PROJECT_META = "project_match_data_260612.pkl"                   # 과제수행기관/논문/특허 성과 조인
MODEL_DIR = "pro-sroberta"

TOPK = 100             # SBERT(과제 임베딩) 코사인 상위 N (LLM 평가 대상)
FINAL = 5              # 최종 추천 수 (상세 근거 생성 대상)
KW_MIN, KW_MAX = 10, 20    # 추출 키워드 개수 범위
DESC_OUT = 30000       # 엑셀 셀 길이 한계(32767) 회피용 출력 설명문 절단
SEP = ";"
# -------------------------------------------------------------------------


def norm_name(s):
    """기업명 정규화 — 법인격 표기/공백 제거 후 소문자화 (match_compa_list 와 동일)."""
    return mcl.norm_name(s)


def cell(rec, col):
    return mcl.cell(rec, col)


# ---- LLM 키워드 추출 ------------------------------------------------------
_KW_SYS = ("당신은 기술 문서에서 핵심 기술 키워드를 추출하는 전문가입니다. "
           "기업의 기술 수요 설명에서 검색·매칭에 유용한 구체적 기술 용어만 선별합니다.")

_KW_GUIDE = (
    "위 [기업 기술 수요]의 수요기술명·수요기술 내용·수요기술 사양에서 핵심 기술 키워드를 "
    f"{KW_MIN}~{KW_MAX}개 추출하세요.\n"
    "  - 소재/공정/성분/기능/대상 등 구체적 기술 명사·전문용어 위주(한글·영문 혼용 허용)\n"
    "  - 일반어(기술, 제품, 개발, 공정, 기반, 적용 등)·수식어·단위·수치는 제외\n"
    "  - 핵심 영문 전문용어(예: Bioavailability, Platycodin D)는 그대로 보존\n"
    "  - 가능한 한 단일/복합 명사 형태의 짧은 키워드로\n"
    "반드시 아래 JSON 형식으로만 답하세요:\n"
    '{"keywords": ["키워드1", "키워드2", ...]}'
)


def _kw_block(d):
    return ("[기업 기술 수요]\n"
            f"- 수요기술명: {d.get('수요기술명','')}\n"
            f"- 수요기술 내용: {d.get('수요기술 내용','')}\n"
            f"- 수요기술 사양: {d.get('수요기술 사양','')}")


def _parse_keywords(text):
    """LLM 출력에서 {\"keywords\": [...]} 를 견고하게 파싱 → 키워드 리스트(중복/공백 제거)."""
    kws = []
    m = re.search(r'\{.*"keywords".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("keywords", []) if isinstance(data, dict) else data
    except Exception:
        # 폴백: "keywords" 배열 내용을 정규식으로 추출
        items = []
        mm = re.search(r'"keywords"\s*:\s*\[(.*?)\]', text, re.DOTALL)
        if mm:
            items = re.findall(r'"([^"]+)"', mm.group(1))
    seen, out = set(), []
    for it in items:
        v = str(it).strip().strip('"').strip()
        k = v.casefold()
        if v and k not in seen:
            seen.add(k)
            out.append(v)
    return out[:KW_MAX]


def extract_keywords(demand):
    """수요기술명/내용/사양 → LLM 으로 핵심 기술 키워드 KW_MIN~KW_MAX 개 추출."""
    user = f"{_kw_block(demand)}\n\n{_KW_GUIDE}"
    msgs = [{"role": "system", "content": _KW_SYS},
            {"role": "user", "content": user}]
    try:
        out = mve.stream_explanation(msgs, max_tokens=600, temperature=0.0, top_p=1.0)
        kws = _parse_keywords(out)
    except Exception as e:
        print(f"      ! 키워드 추출 LLM 실패: {e}")
        kws = []
    # 폴백: LLM 실패/부족 시 규칙기반 명사추출로 보충
    if len(kws) < KW_MIN:
        extra = []
        for col in ("수요기술명", "수요기술 내용", "수요기술 사양"):
            extra += uik.normalize_user_input(demand.get(col, ""))
        seen = {k.casefold() for k in kws}
        for v in extra:
            if v.casefold() not in seen:
                seen.add(v.casefold())
                kws.append(v)
            if len(kws) >= KW_MAX:
                break
    return kws[:KW_MAX]
# -------------------------------------------------------------------------


def build_company_emb_index(cdf):
    """회사 임베딩 DataFrame → {정규화기업명: L2정규화 norm_embed}.
    동일 정규화명이 여럿이면 키워드_리스트가 가장 긴 행을 채택."""
    idx, best_len = {}, {}
    names = cdf["한글업체명"].astype(str).values
    embs = cdf["norm_embed"].values
    kws = cdf["키워드_리스트"].values if "키워드_리스트" in cdf.columns else [None] * len(cdf)
    for i in range(len(cdf)):
        key = norm_name(names[i])
        if not key:
            continue
        klen = len(kws[i]) if isinstance(kws[i], (list, tuple)) else 0
        if key in idx and klen <= best_len.get(key, -1):
            continue
        e = np.asarray(embs[i], dtype=np.float32)
        n = np.linalg.norm(e)
        idx[key] = e / n if n else e
        best_len[key] = klen
    return idx


def build_company_desc_index(cdf):
    """회사 임베딩 DataFrame → {정규화기업명: (기업설명문, desc_ok, desc_issue)}.
    상세 추천 근거에 기업 자체 설명·특성을 반영하기 위한 인덱스.
    동일 정규화명이 여럿이면 설명문이 더 긴(정보 많은) 행을 채택."""
    if "기업설명문" not in cdf.columns:
        return {}
    idx, best = {}, {}
    names = cdf["한글업체명"].astype(str).values
    descs = cdf["기업설명문"].astype(str).values
    oks = cdf["desc_ok"].values if "desc_ok" in cdf.columns else [1] * len(cdf)
    isss = (cdf["desc_issue"].astype(str).values if "desc_issue" in cdf.columns
            else [""] * len(cdf))
    for i in range(len(cdf)):
        key = norm_name(names[i])
        if not key:
            continue
        d = descs[i].strip()
        if d.lower() == "nan":
            d = ""
        if key in idx and len(d) <= best.get(key, -1):
            continue
        try:
            ok = int(oks[i])
        except Exception:
            ok = 1
        iss = str(isss[i]).strip()
        if iss.lower() == "nan":
            iss = ""
        idx[key] = (d, ok, iss)
        best[key] = len(d)
    return idx


def make_encoder(model):
    def encode(text):
        e = model.encode([text], convert_to_numpy=True, normalize_embeddings=False,
                         show_progress_bar=False)[0].astype(np.float32)
        n = np.linalg.norm(e)
        return e / n if n else e
    return encode


def load_corpus():
    """과제/기업 임베딩 + 과제 메타를 1회 로드(여러 담당자에 재사용)."""
    print("  · 과제 임베딩 로드…", flush=True)
    pdf = pd.read_pickle(PROJECT_EMB)
    corpus = {
        "pid": pdf["과제고유번호"].astype(str).values,
        "pname": pdf["과제명"].astype(str).values,
        "pdesc": pdf["과제설명문"].astype(str).values,
        "pkw": pdf["키워드_리스트"].values,
        "promise": pdf["유망성점수"].values.astype(np.float32),
    }
    M = np.vstack([np.asarray(e, dtype=np.float32) for e in pdf["norm_embed"].values])
    M /= np.linalg.norm(M, axis=1, keepdims=True)
    corpus["M"] = M
    del pdf
    print("  · 기업 설명문 로드(상세근거용)…", flush=True)
    cdf = pd.read_pickle(COMPANY_EMB)
    corpus["cdesc_idx"] = build_company_desc_index(cdf)   # 상세 근거용 기업설명문
    del cdf
    print("  · 과제 메타 로드…", flush=True)
    with open(PROJECT_META, "rb") as f:
        pmeta = pickle.load(f)
    corpus["pmeta"] = pmeta
    corpus["org"] = np.array([str(pmeta.get(p, {}).get("과제수행기관명", "")) for p in corpus["pid"]])
    return corpus


def extract_keywords_for(assignee, records):
    """담당자 records 키워드 추출(담당자별 체크포인트) → kw_ckpt 반환 + 키워드 xlsx 저장."""
    ckpt_path = kw_ckpt_path(assignee)
    kw_ckpt = {}
    if os.path.exists(ckpt_path):
        with open(ckpt_path, encoding="utf-8") as f:
            kw_ckpt = json.load(f)
        print(f"      키워드 체크포인트 로드: {len(kw_ckpt)}건")
    kw_rows = []
    for n, rec in enumerate(records, 1):
        demand_no = cell(rec, "번호")
        company = cell(rec, "기업명")
        demand = {c: cell(rec, c) for c in
                  ("수요기술명", "수요기술 내용", "수요기술 사양", "예상 적용 제품 및 서비스")}
        ck = str(demand_no)
        if ck in kw_ckpt:
            kws = kw_ckpt[ck]
        else:
            kws = extract_keywords(demand)
            kw_ckpt[ck] = kws
            with open(ckpt_path, "w", encoding="utf-8") as f:
                json.dump(kw_ckpt, f, ensure_ascii=False)
        row = dict(rec)
        row["키워드"] = SEP.join(kws)
        kw_rows.append(row)
        print(f"      [{demand_no}] {company}: 키워드 {len(kws)}개 [{n}/{len(records)}]")
    out = kw_xlsx_path(assignee)
    pd.DataFrame(kw_rows).to_excel(out, index=False)
    print(f"      → '{out}' 저장({len(kw_rows)}행)")
    return kw_ckpt


def match_for(assignee, records, kw_ckpt, corpus, encode, args):
    """담당자 records 2단계 매칭 + LLM 적합도 + 상세근거(담당자별 체크포인트) → 최종 xlsx 저장."""
    topk, final = args.topk, args.final
    pid, pname, pdesc = corpus["pid"], corpus["pname"], corpus["pdesc"]
    pkw, promise, M = corpus["pkw"], corpus["promise"], corpus["M"]
    pmeta, org = corpus["pmeta"], corpus["org"]
    cdesc_idx = corpus.get("cdesc_idx", {})

    llm_path, ex_path = llm_ckpt_path(assignee), ex_ckpt_path(assignee)
    llm_ckpt = {}
    if os.path.exists(llm_path):
        with open(llm_path, encoding="utf-8") as f:
            llm_ckpt = json.load(f)
        print(f"      LLM점수 체크포인트 로드: {len(llm_ckpt)}건")
    ex_ckpt = {}
    if os.path.exists(ex_path):
        with open(ex_path, encoding="utf-8") as f:
            ex_ckpt = json.load(f)
        print(f"      상세근거 체크포인트 로드: {len(ex_ckpt)}건")

    final_rows = []
    for n, rec in enumerate(records, 1):
        demand_no = cell(rec, "번호")
        company = cell(rec, "기업명")
        kws = kw_ckpt[str(demand_no)]
        kw_string = SEP.join(kws)

        # query 임베딩 (키워드 문자열)
        q = encode(kw_string)

        # 단일 SBERT 매칭: 과제 임베딩 코사인 상위 topk (= LLM 평가 대상)
        cos1 = np.clip(M @ q, -1, 1)
        cand_idx = np.argsort(-cos1)[:topk]

        # LLM 적합도 정량 평가 (수요기술명/내용/사양 충족 0~100) — match_compa_list.llm_score 재사용
        demand = {c: cell(rec, c) for c in
                  ("수요기술명", "수요기술 내용", "수요기술 사양", "예상 적용 제품 및 서비스")}
        demand["기업명"] = company
        demand["keywords"] = kws
        # 상세 근거 생성용 demand: 회사 임베딩 DB의 기업설명문(사업영역·특성)을 추가(있을 때).
        # LLM 점수(llm_score)는 수요기술 기준 그대로 두고, 설명에만 기업 컨텍스트를 반영한다.
        demand_ex = dict(demand)
        cdesc = cdesc_idx.get(norm_name(company))
        if cdesc and cdesc[0]:
            demand_ex["기업설명문"] = cdesc[0]
            demand_ex["desc_ok"] = cdesc[1]
            demand_ex["desc_issue"] = cdesc[2]
        cands = [{"id": j + 1, "과제명": pname[i], "과제설명문": pdesc[i],
                  "키워드": list(pkw[i]) if isinstance(pkw[i], (list, tuple)) else []}
                 for j, i in enumerate(cand_idx)]
        ck = str(demand_no)
        if ck in llm_ckpt:
            id2 = {int(k): tuple(v) for k, v in llm_ckpt[ck].items()}
        else:
            id2 = mcl.llm_score(demand, cands)
            llm_ckpt[ck] = {str(k): list(v) for k, v in id2.items()}
            with open(llm_path, "w", encoding="utf-8") as f:
                json.dump(llm_ckpt, f, ensure_ascii=False)

        # LLM 점수 내림차순 → 최종 final 개 (동점 시 과제 코사인)
        scored = []
        for j, i in enumerate(cand_idx):
            s, reason = id2.get(j + 1, (0, ""))
            scored.append((s, float(cos1[i]), j, i, reason))
        scored.sort(key=lambda x: (-x[0], -x[1]))

        # 과제명 중복 제거: 동일 과제명이 여럿이면 정렬 상 먼저 오는(=점수가 더 높은) 것만
        # 남기고 나머지는 건너뛴다. 중복을 제거하면서 final 개를 채울 때까지 후보를 훑는다.
        seen_titles, selected = set(), []
        for tup in scored:
            tkey = str(pname[tup[3]]).strip().casefold()
            if tkey in seen_titles:
                continue
            seen_titles.add(tkey)
            selected.append(tup)
            if len(selected) >= final:
                break

        for rank, (s, _c1, j, i, reason) in enumerate(selected, 1):
            # 상세 근거: 최종 final 개에 한해서만 생성 (--no-explain 이면 생략).
            # 매칭/점수는 27B로 먼저 끝내고, 상세근거는 별도 패스(35B-A3B 등)에서 채우는
            # 순차 실행을 지원한다. 후보/점수는 체크포인트에 있어 그대로 재사용된다.
            ekey = f"{demand_no}::{pid[i]}"
            if not args.no_explain and ekey not in ex_ckpt:
                meta = pmeta.get(pid[i], {}) if isinstance(pmeta, dict) else {}
                proj = {
                    "pid": pid[i], "과제명": pname[i], "설명": pdesc[i],
                    "유망성": round(float(promise[i]), 1), "수행기관": org[i],
                    "키워드": list(pkw[i]) if isinstance(pkw[i], (list, tuple)) else [],
                    "논문명": meta.get("논문명_리스트") or [],
                    "특허명": meta.get("특허명_리스트") or [],
                    "논문건수": meta.get("논문건수", 0), "특허건수": meta.get("특허건수", 0),
                    "총연구비_상위비율": meta.get("총연구비_상위비율"),
                    "논문건수_상위비율": meta.get("논문건수_상위비율"),
                    "특허건수_상위비율": meta.get("특허건수_상위비율"),
                }
                ex_ckpt[ekey] = mcl.gen_explanation(demand_ex, proj)[:DESC_OUT]
                with open(ex_path, "w", encoding="utf-8") as f:
                    json.dump(ex_ckpt, f, ensure_ascii=False)
            final_rows.append({
                "번호": demand_no, "기업명": company,
                "기술번호": cell(rec, "기술번호"), "수요기술명": cell(rec, "수요기술명"),
                "키워드": kw_string, "rank": rank,
                "LLM점수": s, "LLM판단근거": reason,
                "과제고유번호": pid[i], "과제명": pname[i], "과제수행기관": org[i],
                "유사도_과제코사인": round(float(cos1[i]), 6),
                "유망성점수": round(float(promise[i]), 4),
                "추천근거_상세": ex_ckpt.get(ekey, ""),
                "과제설명문": pdesc[i][:DESC_OUT],
            })
        best = scored[0][0] if scored else 0
        print(f"      [{demand_no}] {company}: SBERT top{len(cand_idx)} "
              f"→ LLM best={best} → 최종 {len(selected)} [{n}/{len(records)}]")

    out = final_xlsx_path(assignee)
    pd.DataFrame(final_rows).to_excel(out, index=False)
    print(f"      → '{out}' 저장(최종추천 {len(final_rows)}행)")

    # 표준 산출물 자동 생성: 15컬럼 정리본(_보완) + 기업별 탭·멀티라인(_보완_기업별)
    try:
        import postprocess_compa as pp
        flat, tabs = pp.make_deliverables(assignee, src=out)
        print(f"      → 후처리: '{flat}', '{tabs}' 생성")
    except Exception as e:
        print(f"      ! 후처리(정리/탭 분리) 실패: {e}")


def resolve_assignees(xl, acol, args):
    """처리할 담당자 목록 결정. --all(시트의 모든 담당자) / --assignee(단일) / 기본값."""
    present = []
    for v in xl[acol].astype(str).str.strip():
        if v and v.lower() != "nan" and v not in present:
            present.append(v)
    if args.all:
        targets = present
    elif args.assignee:
        targets = [a.strip() for a in args.assignee.split(",") if a.strip()]
    else:
        targets = [DEFAULT_ASSIGNEE]
    exclude = {e.strip() for e in (args.exclude or "").split(",") if e.strip()}
    return [a for a in targets if a not in exclude]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--assignee", default="", help="대상 담당자(쉼표로 여러 명). 미지정+--all 아니면 기본='이중연'")
    ap.add_argument("--all", action="store_true", help="시트의 모든 담당자 처리(담당자별 파일 분리)")
    ap.add_argument("--exclude", default="", help="제외할 담당자(쉼표 구분). 예: --all --exclude 이중연")
    ap.add_argument("--limit", type=int, default=0, help="담당자별 처리 수요 수 제한(0=전체, 테스트용)")
    ap.add_argument("--topk", type=int, default=TOPK, help="SBERT 후보 수(LLM 평가 대상)")
    ap.add_argument("--final", type=int, default=FINAL, help="최종 추천 수")
    ap.add_argument("--no-explain", action="store_true",
                    help="상세근거 생성 생략(매칭/LLM점수까지만). 상세근거는 별도 패스에서 "
                         "다른 모델(예: MV_MLX_MODEL=...Qwen3.5-35B-A3B-4bit)로 채운다.")
    ap.add_argument("--keywords-only", action="store_true",
                    help="키워드 추출 + 키워드 xlsx 저장까지만 수행(매칭/LLM 스킵)")
    args = ap.parse_args()

    print("[1/4] 시트 로드 + 담당자 결정…")
    xl = pd.read_excel(XLSX, header=HEADER_ROW)
    acol = xl.columns[ASSIGNEE_COL_IDX]
    assignees = resolve_assignees(xl, acol, args)
    print(f"      대상 담당자({len(assignees)}): {assignees} (대상 컬럼='{acol}', 시트 전체 {len(xl)})")

    print("[2/4] pro-sroberta 로드…")
    model = SentenceTransformer(MODEL_DIR, device="cpu")
    model.eval()
    encode = make_encoder(model)

    corpus = None
    if not args.keywords_only:
        print("[3/4] 과제/기업 임베딩 코퍼스 로드(1회)…")
        corpus = load_corpus()
    else:
        print("[3/4] (keywords-only: 코퍼스 로드 생략)")

    print("[4/4] 담당자별 처리…")
    for ai, assignee in enumerate(assignees, 1):
        target = xl[xl[acol].astype(str).str.strip() == assignee].reset_index(drop=True)
        records = target.to_dict("records")
        if args.limit:
            records = records[:args.limit]
        print(f"  === [{ai}/{len(assignees)}] 담당자='{assignee}' 수요 {len(records)}건 ===")
        kw_ckpt = extract_keywords_for(assignee, records)
        if args.keywords_only:
            continue
        match_for(assignee, records, kw_ckpt, corpus, encode, args)
    print("완료.")


if __name__ == "__main__":
    main()
