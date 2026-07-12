# -*- coding: utf-8 -*-
"""매칭 과제(pid)들의 7개 상세 필드를 apollo 3파일에서 추출 → pid_fields.json.
또 (과제명,수행기관)->pid 매핑도 저장."""
import pandas as pd, pickle, json, glob, re, os
def norm(s): return re.sub(r"\s+", "", str(s)).strip()
SCRATCH = os.environ.get("COMPA_SCRATCH", "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
os.makedirs(SCRATCH, exist_ok=True)
AP = os.environ.get("COMPA_APOLLO_DIR", "/Users/osung/work/apollo")

# --- 매칭 pid + (기업,수요,과제명)->pid  (수행기관 키는 동일과제명 충돌 유발 → 수요 맥락 포함) ---
name2pid = {}
pids = set()
for f in glob.glob("COMPA_*_최종추천.pkl"):
    if "전체" in f: continue
    for r in pickle.load(open(f, "rb")).to_dict("records"):
        pid = str(r["과제고유번호"]); pids.add(pid)
        name2pid.setdefault((norm(r["기업명"]), norm(r["수요기술명"]), norm(r["과제명"])), pid)
for k, e in json.load(open("COMPA_통합best.json")).items():
    for t in e.get("top5", []):
        pid = str(t.get("과제고유번호", ""));
        if pid: pids.add(pid)
        name2pid.setdefault((norm(e["기업명"]), norm(e["수요기술명"]), norm(t["과제명"])), pid)
for t in json.load(open(f"{SCRATCH}/regen_targets.json"))["regen"]:
    pid = str(t["pid"]); pids.add(pid)
    name2pid.setdefault((norm(t["기업명"]), norm(t["수요기술명"]), norm(t["과제명"])), pid)
print("매칭 pid:", len(pids), "| name2pid:", len(name2pid))

# --- dataset_260602: 총연구비/기간/표준분류중/수행기관 ---
ds = pd.read_pickle(f"{AP}/df_project_dataset_260602.pkl")
ds["과제고유번호"] = ds["과제고유번호"].astype(str)
ds = ds[ds["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호")
print("dataset matched:", len(ds))

# --- embeddings_260708: 연구수행주체/연구개발단계 ---
em = pd.read_pickle(f"{AP}/public_RnD_embeddings_pro_260601_with_desc_260708.pkl")
em["과제고유번호"] = em["과제고유번호"].astype(str)
em = em[em["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호")
print("embeddings260708 matched:", len(em))

# --- all_bizno: 표준분류대 ---
bz = pd.read_pickle(f"{AP}/df_project_all_bizno_260310.pkl")
bz["과제고유번호"] = bz["과제고유번호"].astype(str)
bz = bz[bz["과제고유번호"].isin(pids)].drop_duplicates("과제고유번호").set_index("과제고유번호")
print("all_bizno matched:", len(bz))

def g(df, pid, col):
    if pid in df.index and col in df.columns:
        v = df.loc[pid, col]
        s = str(v).strip()
        return "" if s in ("nan", "None", "") else s
    return ""

fields = {}
for pid in pids:
    fields[pid] = {
        "과제고유번호": pid,
        "연구시작일": g(ds, pid, "연구시작일"),
        "연구종료일": g(ds, pid, "연구종료일"),
        "총연구기간": g(ds, pid, "총연구기간"),
        "총연구비": g(ds, pid, "총연구비"),
        "표준분류중": g(ds, pid, "과학기술표준분류1-중"),
        "표준분류대": g(bz, pid, "과학기술표준분류1-대"),
        "과제수행기관명": g(ds, pid, "과제수행기관명"),
        "연구수행주체": g(em, pid, "연구수행주체"),
        "연구개발단계": g(em, pid, "연구개발단계"),
    }

json.dump({"name2pid": {f"{a}||{b}||{c}": p for (a, b, c), p in name2pid.items()}, "fields": fields},
          open(f"{SCRATCH}/pid_fields.json", "w"), ensure_ascii=False)
print("saved pid_fields.json")
# 커버리지 요약
import collections
c = collections.Counter()
for v in fields.values():
    for k in ("총연구비", "표준분류중", "표준분류대", "연구수행주체", "연구개발단계", "연구시작일"):
        if v[k]: c[k] += 1
print("coverage:", dict(c), "of", len(fields))
