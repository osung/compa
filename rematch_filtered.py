# -*- coding: utf-8 -*-
"""필터(연구수행주체 4종 ∧ 제출년도 최근5년=2020~) 적용 재매칭.
지정 임베딩 파일(public_RnD_embeddings_pro_with_desc_260708.pkl)로 코퍼스를 필터링한 뒤
compa_match.match_for 로 재검색·35B 재채점·Top5·근거 생성. 기존 키워드 캐시 재사용.

사용: python rematch_filtered.py --n 10 [--start 1] [--tag 필터10]
"""
import argparse, glob, json, os, pickle, sys, time
import numpy as np, pandas as pd
from sentence_transformers import SentenceTransformer
import compa_match as cm

AP = os.environ.get("COMPA_APOLLO_DIR", "/Users/osung/work/apollo")
EMB_FILE = f"{AP}/public_RnD_embeddings_pro_with_desc_260708.pkl"
ALLOW = {"대학", "출연연구소", "국공립연구소", "정부부처"}
YEAR_MIN = 2020            # 제출년도 최근 5년(2020~2024)

def log(*a): print(*a, flush=True)

def build_filtered_corpus():
    t0 = time.time()
    log(f"· 임베딩 로드+필터 (주체 {sorted(ALLOW)} ∧ 제출년도>={YEAR_MIN})…")
    pdf = pd.read_pickle(EMB_FILE)
    y = pd.to_numeric(pdf["제출년도"], errors="coerce")
    mask = pdf["연구수행주체"].isin(ALLOW) & (y >= YEAR_MIN)
    pdf = pdf[mask].reset_index(drop=True)
    log(f"  필터 통과 {len(pdf)}건 ({time.time()-t0:.0f}s)")
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
    log("· BM25 색인(필터 코퍼스)…")
    corpus["bm25"] = cm.build_bm25(corpus["pname"], corpus["pdesc"], corpus["pkw"])
    log("· 기업 설명문 로드…")
    corpus["cdesc_idx"] = cm.build_company_desc_index(pd.read_pickle(cm.COMPANY_EMB))
    log("· 과제 메타 로드…")
    with open(cm.PROJECT_META, "rb") as f:
        pmeta = pickle.load(f)
    corpus["pmeta"] = pmeta
    corpus["org"] = np.array([str(pmeta.get(p, {}).get("과제수행기관명", "")) for p in corpus["pid"]])
    log(f"  코퍼스 준비 완료 ({time.time()-t0:.0f}s)")
    return corpus

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=10)
    ap.add_argument("--start", type=int, default=1)
    ap.add_argument("--tag", default="필터10")
    a = ap.parse_args()

    xl = pd.read_excel(cm.XLSX, header=cm.HEADER_ROW)
    recs = xl.to_dict("records")
    recs = [r for r in recs if str(cm.cell(r, "번호")).strip() not in ("", "nan", "None")]
    recs = recs[a.start - 1: a.start - 1 + a.n]
    log(f"대상 수요 {len(recs)}건: 번호 {[cm.cell(r,'번호') for r in recs]}")

    # 키워드 캐시(기존) → tag 용 kw_ckpt
    kwmap = {}
    for f in glob.glob("compa2stage_*_keywords_ckpt.json"):
        for k, v in json.load(open(f)).items():
            kwmap[int(float(k))] = v
    kw_ckpt = {}
    for r in recs:
        dn = str(cm.cell(r, "번호"))
        kw_ckpt[dn] = kwmap.get(int(float(dn)), [])
    with open(cm.kw_ckpt_path(a.tag), "w", encoding="utf-8") as f:
        json.dump(kw_ckpt, f, ensure_ascii=False)

    log("· pro-sroberta 로드…")
    model = SentenceTransformer(cm.MODEL_DIR, device="cpu"); model.eval()
    encode = cm.make_encoder(model)

    corpus = build_filtered_corpus()

    log("· 35B 로드…")
    cm.load_model_blocking(progress_cb=lambda m: log("  " + m))

    args = argparse.Namespace(topk=cm.TOPK, final=cm.FINAL, no_explain=False)
    t0 = time.time()
    cm.match_for(a.tag, recs, kw_ckpt, corpus, encode, args)
    log(f"\n✔ 재매칭 완료 ({len(recs)}수요, {time.time()-t0:.0f}s) → COMPA_{a.tag}_최종추천.xlsx/.pkl")

if __name__ == "__main__":
    main()
