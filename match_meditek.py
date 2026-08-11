# -*- coding: utf-8 -*-
"""MEDITEK 수요기술 매칭 — build_meditek_demands.py 산출 DF → Top5 추천 + 근거.

코퍼스 필터는 기존 78수요 최종본과 동일(연구수행주체 4종 ∧ 제출년도 2020~).
키워드 체크포인트는 tag 로 분리되므로 기존 COMPA 번호(1~78)와 충돌하지 않는다.
※ 기존 키워드 캐시를 번호로 재사용하면 안 된다(같은 번호가 다른 기업임).

사용: python match_meditek.py [--tag MEDITEK] [--limit 0]
산출: COMPA_<tag>_최종추천.xlsx/.pkl, COMPA_<tag>_키워드.xlsx
"""
import argparse
import time

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer

import compa_match as cm
import meditek_segment as ms
from rematch_filtered import build_filtered_corpus

DEMANDS_PKL = "MEDITEK_수요기술.pkl"


def log(*a):
    print(*a, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=DEMANDS_PKL)
    ap.add_argument("--tag", default="MEDITEK")
    ap.add_argument("--limit", type=int, default=0, help="처리 수요 수 제한(0=전체)")
    ap.add_argument("--only", default="", help="처리할 번호(쉼표 구분). 예: --only 3,8,10,12")
    ap.add_argument("--demand-scope", default="all", choices=sorted(ms.SCOPES),
                    help="선정 근거에 포함할 원문 구성요소: "
                         "tech=기술수요만, tech+intro=기술수요+회사소개, all=업종까지(기본)")
    ap.add_argument("--topk", type=int, default=cm.TOPK)
    ap.add_argument("--final", type=int, default=cm.FINAL)
    ap.add_argument("--dedupe-phase", action="store_true",
                    help="같은 기관의 연차 후속과제(…Ⅵ/…Ⅶ, N차)를 한 건으로 보고 Top5 선정")
    a = ap.parse_args()

    df = pd.read_pickle(a.src)
    skipped = df[~df["매칭가능"]]
    df = df[df["매칭가능"]].reset_index(drop=True)
    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        df = df[df["번호"].isin(keep)].reset_index(drop=True)
    if a.limit:
        df = df.head(a.limit)

    # 선정 근거 재구성: tech / tech+intro 는 세그먼트에서 해당 부분만 사용
    if a.demand_scope != "all":
        if "세그_기술수요" not in df.columns:
            raise SystemExit("세그먼트 컬럼이 없습니다. "
                             "먼저 `python build_meditek_demands.py --segment` 실행 필요")
        dropped = []
        for i, r in df.iterrows():
            seg = {lab: (r.get(f"세그_{lab}") or "").split("\n") for lab in ms.LABELS}
            seg = {k: [x for x in v if x.strip()] for k, v in seg.items()}
            text = ms.compose(seg, a.demand_scope)
            if not text:
                dropped.append((r["번호"], r["기업명"]))
            df.at[i, "수요기술 내용"] = text
        if dropped:
            log(f"! 선정 근거 비어 제외 {len(dropped)}건("
                f"scope={a.demand_scope}): {[f'{n}:{c}' for n, c in dropped]}")
            df = df[df["수요기술 내용"].str.len() > 0].reset_index(drop=True)

    recs = df.to_dict("records")
    log(f"선정 근거 구성(scope): {a.demand_scope} → {ms.SCOPES[a.demand_scope]}")
    log(f"대상 수요 {len(recs)}건: {[r['번호'] for r in recs]}")
    if len(skipped):
        log(f"제외(수요기술 내용 없음) {len(skipped)}건: {list(skipped['번호'])}")

    log("· 키워드 추출(35B)…")
    cm.load_model_blocking(progress_cb=lambda m: log("  " + m))
    kw_ckpt = cm.extract_keywords_for(a.tag, recs)

    log("· pro-sroberta 로드…")
    model = SentenceTransformer(cm.MODEL_DIR, device="cpu")
    model.eval()
    encode = cm.make_encoder(model)

    corpus = build_filtered_corpus()

    args = argparse.Namespace(topk=a.topk, final=a.final, no_explain=False,
                              dedupe_phase=a.dedupe_phase)
    t0 = time.time()
    cm.match_for(a.tag, recs, kw_ckpt, corpus, encode, args)
    log(f"\n✔ 매칭 완료 ({len(recs)}수요, {time.time()-t0:.0f}s) "
        f"→ COMPA_{a.tag}_최종추천.xlsx/.pkl")


if __name__ == "__main__":
    main()
