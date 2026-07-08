# -*- coding: utf-8 -*-
"""45건 상세근거(4섹션) + 짧은 판단근거를 로컬 35B(MLX)로 재생성. 항목별 체크포인트."""
import json, os, pickle, sys, time
import numpy as np
import pandas as pd

SCRATCH = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"
OUT = f"{SCRATCH}/regen_out.json"

import compa_match as cm

def log(*a):
    print(*a, flush=True)

tgt = json.load(open(f"{SCRATCH}/regen_targets.json"))["regen"]
log(f"targets: {len(tgt)}")

pids = sorted({t["pid"] for t in tgt})
log(f"unique pids: {len(pids)}")

# --- 과제 메타(pmeta) ---
log("loading pmeta …"); t0 = time.time()
with open(cm.PROJECT_META, "rb") as f:
    pmeta = pickle.load(f)
log(f"  pmeta {len(pmeta)} ({time.time()-t0:.0f}s)")

# --- PROJECT_EMB: pid -> (설명, 유망성, 키워드) (45 pid만 추출 후 폐기) ---
log("loading PROJECT_EMB …"); t0 = time.time()
pdf = pd.read_pickle(cm.PROJECT_EMB)
pdf["과제고유번호"] = pdf["과제고유번호"].astype(str)
sub = pdf[pdf["과제고유번호"].isin(pids)]
pemb = {}
for r in sub.to_dict("records"):
    pemb[str(r["과제고유번호"])] = (
        str(r.get("과제설명문") or ""),
        float(r.get("유망성점수") or 0.0),
        list(r.get("키워드_리스트") or []),
    )
del pdf, sub
log(f"  PROJECT_EMB matched {len(pemb)}/{len(pids)} ({time.time()-t0:.0f}s)")

# --- COMPANY_EMB: 기업설명문 인덱스 ---
log("loading COMPANY_EMB …"); t0 = time.time()
cdf = pd.read_pickle(cm.COMPANY_EMB)
cdesc_idx = cm.build_company_desc_index(cdf)
del cdf
log(f"  cdesc_idx {len(cdesc_idx)} ({time.time()-t0:.0f}s)")

# --- 재개용 체크포인트 로드 ---
done = {}
if os.path.exists(OUT):
    done = json.load(open(OUT))
    log(f"resume: {len(done)} already done")

def keyof(t): return f"{t['번호']}::{t['pid']}"

# --- 모델 로드 ---
log("loading model (35B MLX) …"); t0 = time.time()
cm.load_model_blocking(progress_cb=lambda m: log("  " + m))
log(f"  model ready ({time.time()-t0:.0f}s)")

for n, t in enumerate(tgt, 1):
    k = keyof(t)
    if k in done and done[k].get("추천근거_상세"):
        log(f"[{n}/{len(tgt)}] skip {k}"); continue
    demand = {
        "수요기술명": t["수요기술명"],
        "수요기술 내용": t["수요기술 내용"],
        "수요기술 사양": t["수요기술 사양"],
        "예상 적용 제품 및 서비스": "",
        "기업명": t["기업명"],
        "keywords": t["keywords"],
    }
    demand_ex = dict(demand)
    cd = cdesc_idx.get(cm.norm_name(t["기업명"]))
    if cd and cd[0]:
        demand_ex["기업설명문"], demand_ex["desc_ok"], demand_ex["desc_issue"] = cd

    desc, promise, pkw = pemb.get(t["pid"], (t.get("과제설명문_fallback", ""), 0.0, []))
    meta = pmeta.get(t["pid"], {})
    proj = {
        "pid": t["pid"], "과제명": t["과제명"], "설명": desc,
        "유망성": round(float(promise), 1),
        "수행기관": meta.get("과제수행기관명") or t.get("수행기관", ""),
        "키워드": list(pkw),
        "논문명": meta.get("논문명_리스트") or [],
        "특허명": meta.get("특허명_리스트") or [],
        "논문건수": meta.get("논문건수", 0), "특허건수": meta.get("특허건수", 0),
        "총연구비_상위비율": meta.get("총연구비_상위비율"),
        "논문건수_상위비율": meta.get("논문건수_상위비율"),
        "특허건수_상위비율": meta.get("특허건수_상위비율"),
    }

    ts = time.time()
    detail = cm.gen_explanation(demand_ex, proj)[:cm.DESC_OUT]
    # 짧은 판단근거: 단건 llm_score
    cand = [{"id": 1, "과제명": t["과제명"], "과제설명문": desc or t.get("과제설명문_fallback", ""),
             "키워드": list(pkw)}]
    try:
        sc = cm.llm_score(demand, cand)
        score, reason = sc.get(1, (0, ""))
    except Exception as e:
        score, reason = 0, ""
        log(f"    ! llm_score fail: {e}")

    done[k] = {"기업명": t["기업명"], "수요기술명": t["수요기술명"], "과제명": t["과제명"],
               "수행기관": proj["수행기관"], "pid": t["pid"],
               "판단근거": reason, "LLM점수": score, "추천근거_상세": detail}
    json.dump(done, open(OUT, "w"), ensure_ascii=False, indent=1)
    log(f"[{n}/{len(tgt)}] {k} done ({time.time()-ts:.0f}s) score={score} len={len(detail)}")

log("ALL DONE:", len(done))
