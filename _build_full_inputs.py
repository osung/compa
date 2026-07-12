# -*- coding: utf-8 -*-
"""78수요 전체 보고서 입력 생성(필터 재매칭 3배치 통합):
 - scratchpad/demand_field.json          (번호→6T, XLSX)
 - COMPA_통합best.json                     (78수요 Top5+근거)  ← 정본 갱신
 - scratchpad/pid_fields.json             (pid→단계·주체·연구책임자명·국가연구자번호)
백업: 기존 COMPA_통합best.json → scratchpad/정본백업/
"""
import json, os, shutil, glob
import pandas as pd
import compa_match as cm

SCRATCH = os.environ.get("COMPA_SCRATCH", "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
os.makedirs(SCRATCH, exist_ok=True)
AP = os.environ.get("COMPA_APOLLO_DIR", "/Users/osung/work/apollo")
BATCHES = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]

# --- 1) XLSX 수요 메타 + 6T ---
xl = pd.read_excel(cm.XLSX, header=cm.HEADER_ROW)
field6t, dmeta = {}, {}
for r in xl.to_dict("records"):
    no = str(cm.cell(r, "번호")).strip()
    if no in ("", "nan", "None"): continue
    field6t[no] = str(cm.cell(r, "수요기술 분야(6T)")).strip()
    dmeta[no] = {"기업명": str(cm.cell(r, "기업명")), "수요기술명": str(cm.cell(r, "수요기술명")),
                 "수요기술 내용": str(cm.cell(r, "수요기술 내용") or ""),
                 "수요기술 사양": str(cm.cell(r, "수요기술 사양") or "")}
json.dump(field6t, open(f"{SCRATCH}/demand_field.json", "w"), ensure_ascii=False)
print("demand_field.json:", len(field6t))

# --- 2) 3배치 통합 → 통합best.json ---
df = pd.concat([pd.read_pickle(f) for f in BATCHES], ignore_index=True)
df["번호"] = df["번호"].astype(str)
df["과제고유번호"] = df["과제고유번호"].astype(str)
best = {}
for no in sorted(df["번호"].unique(), key=int):
    g = df[df["번호"] == no].sort_values("rank")
    dm = dmeta.get(no, {})
    top5 = [{"rank": int(r["rank"]), "과제고유번호": str(r["과제고유번호"]), "과제명": r["과제명"],
             "수행기관": r["과제수행기관"], "LLM점수": int(r["LLM점수"]), "판단근거": r["LLM판단근거"],
             "과제설명문": r["과제설명문"], "추천근거_상세": r["추천근거_상세"]}
            for r in g.to_dict("records")]
    best[no] = {"기업명": dm.get("기업명", g.iloc[0]["기업명"]),
                "수요기술명": dm.get("수요기술명", g.iloc[0]["수요기술명"]),
                "수요기술 내용": dm.get("수요기술 내용", ""), "수요기술 사양": dm.get("수요기술 사양", ""),
                "top5": top5}
os.makedirs(f"{SCRATCH}/정본백업", exist_ok=True)
if os.path.exists("COMPA_통합best.json"):
    shutil.copy("COMPA_통합best.json", f"{SCRATCH}/정본백업/COMPA_통합best.json.pre필터")
json.dump(best, open("COMPA_통합best.json", "w"), ensure_ascii=False, indent=1)
print("COMPA_통합best.json:", len(best), "수요")

# --- 3) pid_fields.json (단계·주체·PI) ---
pids = set(df["과제고유번호"])
emb = pd.read_pickle(f"{AP}/public_RnD_embeddings_pro_with_desc_260708.pkl")
emb["과제고유번호"] = emb["과제고유번호"].astype(str)
em = emb[emb["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호"); del emb
pi = pd.read_pickle(f"{AP}/public_RnD_PI_260610.pkl")
pi["과제고유번호"] = pi["과제고유번호"].astype(str)
pim = pi[pi["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호"); del pi

def g(fr, pid, col):
    if pid in fr.index and col in fr.columns:
        s = str(fr.loc[pid, col]).strip()
        return "" if s in ("nan", "None", "") else s
    return ""

fields = {pid: {"연구개발단계": g(em, pid, "연구개발단계"), "연구수행주체": g(em, pid, "연구수행주체"),
                "연구책임자명": g(pim, pid, "연구책임자명"), "국가연구자번호": g(pim, pid, "국가연구자번호")}
          for pid in pids}
json.dump(fields, open(f"{SCRATCH}/pid_fields.json", "w"), ensure_ascii=False)
import collections
c = collections.Counter()
for v in fields.values():
    for k in v:
        if v[k]: c[k] += 1
print("pid_fields.json:", len(fields), "pid | 커버리지:", dict(c))
