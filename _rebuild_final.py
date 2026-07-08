# -*- coding: utf-8 -*-
"""정본(통합best.json + 담당자 pkl)을 docx 최종 Top5 선정으로 교체 재구성.
식별키 = (기업, 수요, 과제명). pid/메타는 소스(pkl→통합best→regen)에서 가져온다(docx pid 불신).
기본 dry-run. --write 시 백업 후 기록."""
import json, pickle, glob, re, sys, shutil, os
import numpy as np, pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

def norm(s): return re.sub(r"\s+", "", str(s)).strip()
SCRATCH = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"
WRITE = "--write" in sys.argv

# ---------- 1) docx 최종 선정(수요→rank·과제명) 추출 ----------
d = Document("COMPA_최종Top5_보고서.docx")
demands = []; cur = None
for ch in list(d.element.body):
    if ch.tag == qn("w:p"):
        p = Paragraph(ch, d); t = p.text.strip(); st = p.style.name if p.style else ""
        if st == "Heading 2" and t.startswith("[수요"):
            cur = {"수요기술명": re.sub(r"^\[수요\s*\d+\]\s*", "", t).strip(), "기업명": None, "entries": []}
            demands.append(cur)
        elif st == "Heading 3" and re.match(r"Top\d", t) and cur is not None:
            m = re.match(r"Top(\d+)\.\s*(.*)", t)
            cur["entries"].append({"rank": int(m.group(1)),
                                   "과제명": re.sub(r"\s*\(출처[:：].*?\)\s*$", "", m.group(2)).strip()})
    elif ch.tag == qn("w:tbl") and cur is not None:
        tb = Table(ch, d)
        if tb.rows[0].cells[0].text.strip() == "기업명" and cur["기업명"] is None:
            cur["기업명"] = tb.rows[0].cells[1].text.strip()
print(f"docx 수요:{len(demands)} | Top:{sum(len(x['entries']) for x in demands)}")

# ---------- 2) 소스 맵 (키=(기업,수요,과제명)) ----------
pk_row = {}; dem_meta = {}; assignee_of = {}
for f in glob.glob("COMPA_*_최종추천.pkl"):
    if "전체" in f: continue
    a = re.search(r"COMPA_(.+?)_최종추천", f).group(1)
    for r in pickle.load(open(f, "rb")).to_dict("records"):
        dk = (norm(r["기업명"]), norm(r["수요기술명"]))
        pk_row[(dk[0], dk[1], norm(r["과제명"]))] = r
        dem_meta.setdefault(dk, {"번호": r["번호"], "기술번호": r["기술번호"], "키워드": r["키워드"],
                                 "기업명": r["기업명"], "수요기술명": r["수요기술명"]})
        assignee_of[dk] = a

jb = json.load(open("COMPA_통합best.json"))
jb_dem = {}; jb_item = {}
for di, e in jb.items():
    dk = (norm(e["기업명"]), norm(e["수요기술명"]))
    jb_dem[dk] = {"수요기술 내용": e.get("수요기술 내용", ""), "수요기술 사양": e.get("수요기술 사양", ""),
                  "기업명": e["기업명"], "수요기술명": e["수요기술명"], "_idx": di}
    for t in e.get("top5", []):
        jb_item[(dk[0], dk[1], norm(t["과제명"]))] = t

ro = {}
for v in json.load(open(f"{SCRATCH}/regen_out.json")).values():
    ro[(norm(v["기업명"]), norm(v["수요기술명"]), norm(v["과제명"]))] = v
rt = {}
for t in json.load(open(f"{SCRATCH}/regen_targets.json"))["regen"]:
    rt[(norm(t["기업명"]), norm(t["수요기술명"]), norm(t["과제명"]))] = t

# ---------- 3) 소스 해석(1차) : pid·src 확정 ----------
def resolve(dk, name):
    k = (dk[0], dk[1], norm(name))
    if k in pk_row: return "pkl", pk_row[k]
    if k in jb_item: return "jb", jb_item[k]
    if k in ro: return "regen", ro[k]
    return "?", None

resolved = []; miss = []; src_cnt = {"pkl": 0, "jb": 0, "regen": 0}
for dm in demands:
    dk = (norm(dm["기업명"]), norm(dm["수요기술명"]))
    for e in dm["entries"]:
        src, row = resolve(dk, e["과제명"])
        if src == "?": miss.append((dm["기업명"], e["과제명"][:35])); continue
        src_cnt[src] += 1
        resolved.append((dk, dm, e, src, row))
print("해석:", len(resolved), "| 출처:", src_cnt, "| 실패:", len(miss))
for x in miss[:10]: print("  MISS", x)

# ---------- 4) apollo 유망성(pkl 외 소스 pid용) ----------
need_prom = set()
for dk, dm, e, src, row in resolved:
    if src != "pkl":
        pid = str(row["과제고유번호"] if src == "jb" else row["pid"])
        need_prom.add(pid)
prom = {}
if need_prom:
    ds = pd.read_pickle("/Users/osung/work/apollo/df_project_dataset_260602.pkl")
    ds["과제고유번호"] = ds["과제고유번호"].astype(str)
    prom = {r["과제고유번호"]: float(r["유망성점수"]) for _, r in
            ds[ds["과제고유번호"].isin(need_prom)][["과제고유번호", "유망성점수"]].iterrows()}
    del ds

# ---------- 5) 행 조립 ----------
def make_row(dk, dm, e, src, row):
    if src == "pkl":
        r = row
        return dict(과제고유번호=str(r["과제고유번호"]), 과제명=r["과제명"], 수행기관=r["과제수행기관"],
                    LLM점수=int(r["LLM점수"]), 판단근거=r["LLM판단근거"], 과제설명문=r["과제설명문"],
                    추천근거_상세=r["추천근거_상세"], 유사도=r.get("유사도_과제코사인"),
                    유망성=r.get("유망성점수"),
                    출처=jb_item.get((dk[0], dk[1], norm(e["과제명"])), {}).get("출처", "기존"))
    if src == "jb":
        t = row; pid = str(t["과제고유번호"])
        return dict(과제고유번호=pid, 과제명=t["과제명"], 수행기관=t.get("수행기관", ""),
                    LLM점수=int(t.get("LLM점수", 0)), 판단근거=t.get("판단근거", ""),
                    과제설명문=t.get("과제설명문", ""), 추천근거_상세=t.get("추천근거_상세", ""),
                    유사도=np.nan, 유망성=prom.get(pid), 출처=t.get("출처", "기존"))
    v = row; pid = str(v["pid"]); t = rt.get((dk[0], dk[1], norm(e["과제명"])), {})
    return dict(과제고유번호=pid, 과제명=v["과제명"], 수행기관=v["수행기관"],
                LLM점수=int(v["LLM점수"]), 판단근거=v.get("매칭근거_short") or v.get("판단근거"),
                과제설명문=t.get("과제설명문_fallback", ""), 추천근거_상세=v["추천근거_상세"],
                유사도=np.nan, 유망성=prom.get(pid), 출처=t.get("출처", "신규"))

# 검증: demand 메타/통합best 수요 커버
nodem = sum(1 for dk, *_ in resolved if dk not in dem_meta)
nojb = sum(1 for dk, *_ in resolved if dk not in jb_dem)
print("demand 메타 없음:", nodem, "| 통합best 수요 없음:", nojb)
# 각 수요 5개인지
from collections import Counter
per = Counter((dm["기업명"], dm["수요기술명"]) for dk, dm, e, s, r in resolved)
bad = [k for k, c in per.items() if c != 5]
print("Top5 아닌 수요:", len(bad), bad[:5])

if not WRITE:
    print("\n[dry-run] 이상 없으면 --write 로 기록.")
    sys.exit(0)

# ---------- 6) 백업 + 기록 ----------
bkdir = f"{SCRATCH}/정본백업"
os.makedirs(bkdir, exist_ok=True)
shutil.copy("COMPA_통합best.json", f"{bkdir}/COMPA_통합best.json")
for f in glob.glob("COMPA_*_최종추천.pkl") + glob.glob("COMPA_*_최종추천.xlsx"):
    if "전체" in f: continue
    shutil.copy(f, f"{bkdir}/{os.path.basename(f)}")
print("백업 →", bkdir)

# 6-1) 통합best.json 재구성(수요별 top5 교체)
new_jb = {}
# 수요 순서 유지: 기존 jb 순서대로
dk2entries = {}
for dk, dm, e, src, row in resolved:
    dk2entries.setdefault(dk, []).append((e["rank"], make_row(dk, dm, e, src, row)))
for di, e in jb.items():
    dk = (norm(e["기업명"]), norm(e["수요기술명"]))
    ents = sorted(dk2entries.get(dk, []), key=lambda x: x[0])
    top5 = []
    for rank, r in ents:
        top5.append({"rank": rank, "과제고유번호": r["과제고유번호"], "과제명": r["과제명"],
                     "수행기관": r["수행기관"], "LLM점수": r["LLM점수"], "출처": r["출처"],
                     "판단근거": r["판단근거"], "과제설명문": r["과제설명문"],
                     "추천근거_상세": r["추천근거_상세"]})
    new_jb[di] = {"기업명": e["기업명"], "수요기술명": e["수요기술명"],
                  "수요기술 내용": e.get("수요기술 내용", ""), "수요기술 사양": e.get("수요기술 사양", ""),
                  "top5": top5}
json.dump(new_jb, open("COMPA_통합best.json", "w"), ensure_ascii=False, indent=1)
print("통합best.json 재작성:", len(new_jb), "수요")

# 6-2) 담당자 pkl 재구성
COLS = ["번호", "기업명", "기술번호", "수요기술명", "키워드", "rank", "LLM점수", "LLM판단근거",
        "과제고유번호", "과제명", "과제수행기관", "유사도_과제코사인", "유망성점수", "추천근거_상세", "과제설명문"]
by_assignee = {}
for dk, dm, e, src, row in resolved:
    a = assignee_of.get(dk)
    meta = dem_meta[dk]; r = make_row(dk, dm, e, src, row)
    by_assignee.setdefault(a, []).append({
        "번호": meta["번호"], "기업명": meta["기업명"], "기술번호": meta["기술번호"],
        "수요기술명": meta["수요기술명"], "키워드": meta["키워드"], "rank": e["rank"],
        "LLM점수": r["LLM점수"], "LLM판단근거": r["판단근거"], "과제고유번호": r["과제고유번호"],
        "과제명": r["과제명"], "과제수행기관": r["수행기관"], "유사도_과제코사인": r["유사도"],
        "유망성점수": r["유망성"], "추천근거_상세": r["추천근거_상세"], "과제설명문": r["과제설명문"]})
for a, rows in by_assignee.items():
    df = pd.DataFrame(rows, columns=COLS).sort_values(["번호", "rank"]).reset_index(drop=True)
    df.to_pickle(f"COMPA_{a}_최종추천.pkl")
    df.to_excel(f"COMPA_{a}_최종추천.xlsx", index=False)
    print(f"  COMPA_{a}_최종추천.pkl/.xlsx: {len(df)}행")
print("완료.")
