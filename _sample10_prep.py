# -*- coding: utf-8 -*-
"""필터10 샘플 보고서 입력 생성:
 - scratchpad/demand_field.json  (번호→6T, XLSX)
 - scratchpad/통합best_필터10.json (10수요 Top5+근거, 필터10 pkl + XLSX 내용/사양)
 - scratchpad/pid_fields.json     (pid→연구개발단계·연구수행주체·연구책임자명·국가연구자번호)
"""
import json, pandas as pd
import compa_match as cm

SCRATCH = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"
AP = "/Users/osung/work/apollo"

# --- 1) demand_field.json (전체 78, 번호→6T) ---
xl = pd.read_excel(cm.XLSX, header=cm.HEADER_ROW)
field6t = {}
demand_meta = {}
for r in xl.to_dict("records"):
    no = str(cm.cell(r, "번호")).strip()
    if no in ("", "nan", "None"): continue
    field6t[no] = str(cm.cell(r, "수요기술 분야(6T)")).strip()
    demand_meta[no] = {
        "기업명": str(cm.cell(r, "기업명")),
        "수요기술명": str(cm.cell(r, "수요기술명")),
        "수요기술 내용": str(cm.cell(r, "수요기술 내용") or ""),
        "수요기술 사양": str(cm.cell(r, "수요기술 사양") or ""),
    }
json.dump(field6t, open(f"{SCRATCH}/demand_field.json", "w"), ensure_ascii=False)
print("demand_field.json:", len(field6t), "건")

# --- 2) 통합best_필터10.json ---
df = pd.read_pickle("COMPA_필터10_최종추천.pkl")
df["번호"] = df["번호"].astype(str)
best = {}
for no, g in df.groupby("번호"):
    dm = demand_meta.get(no, {})
    top5 = []
    for r in g.sort_values("rank").to_dict("records"):
        top5.append({
            "rank": int(r["rank"]), "과제고유번호": str(r["과제고유번호"]),
            "과제명": r["과제명"], "수행기관": r["과제수행기관"],
            "LLM점수": int(r["LLM점수"]), "판단근거": r["LLM판단근거"],
            "과제설명문": r["과제설명문"], "추천근거_상세": r["추천근거_상세"],
        })
    best[no] = {"기업명": dm.get("기업명", g.iloc[0]["기업명"]),
                "수요기술명": dm.get("수요기술명", g.iloc[0]["수요기술명"]),
                "수요기술 내용": dm.get("수요기술 내용", ""),
                "수요기술 사양": dm.get("수요기술 사양", ""),
                "top5": top5}
json.dump(best, open(f"{SCRATCH}/통합best_필터10.json", "w"), ensure_ascii=False, indent=1)
print("통합best_필터10.json:", len(best), "수요")

# --- 3) pid_fields.json (단계·주체·PI) ---
pids = {str(p) for p in df["과제고유번호"]}
emb = pd.read_pickle(f"{AP}/public_RnD_embeddings_pro_with_desc_260708.pkl")
emb["과제고유번호"] = emb["과제고유번호"].astype(str)
em = emb[emb["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호")
del emb
pi = pd.read_pickle(f"{AP}/public_RnD_PI_260610.pkl")
pi["과제고유번호"] = pi["과제고유번호"].astype(str)
pim = pi[pi["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호")
del pi

def g(df, pid, col):
    if pid in df.index and col in df.columns:
        v = df.loc[pid, col]; s = str(v).strip()
        return "" if s in ("nan", "None", "") else s
    return ""

fields = {}
for pid in pids:
    fields[pid] = {
        "연구개발단계": g(em, pid, "연구개발단계"),
        "연구수행주체": g(em, pid, "연구수행주체"),
        "연구책임자명": g(pim, pid, "연구책임자명"),
        "국가연구자번호": g(pim, pid, "국가연구자번호"),
    }
json.dump(fields, open(f"{SCRATCH}/pid_fields.json", "w"), ensure_ascii=False)  # gen_report: 평면 {pid:{...}}
# 커버리지
import collections
c = collections.Counter()
for v in fields.values():
    for k in v:
        if v[k]: c[k] += 1
print("pid_fields.json:", len(fields), "pid | 커버리지:", dict(c))
