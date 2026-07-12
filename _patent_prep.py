# -*- coding: utf-8 -*-
"""매칭 과제(pid)별 특허 실적을 df_pr_patent_260710_detail.pkl 에서 추출 → scratchpad/pid_patents.json.
- 다년도 과제: 해당 pid 의 모든 연도 레코드 포함(데이터가 pid 로 묶여 있어 전 연도 자동 포함).
- 출원/등록 분리 레코드는 (과제, 출원번호)로 병합해 하나의 특허로: 등록정보 우선, 출원정보 병기.
- 정렬: 등록 먼저(등록일 desc), 그다음 출원-only(출원일 desc)."""
import json, os, pandas as pd

SCRATCH = os.environ.get("COMPA_SCRATCH", "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
os.makedirs(SCRATCH, exist_ok=True)
PF = os.path.join(os.environ.get("COMPA_APOLLO_DIR", "/Users/osung/work/apollo"), "df_pr_patent_260710_detail.pkl")

def s(v):
    v = str(v).strip()
    return "" if v in ("nan", "None", "NaT") else v

# 출원/등록 국가코드 → 국가명 (XI=PCT 'PCT/KR…', XU=유럽 EP 출원번호 형식)
_COUNTRY = {"KR": "한국", "US": "미국", "CN": "중국", "JP": "일본",
            "EP": "유럽", "XU": "유럽", "XI": "PCT", "WO": "PCT"}
def country(code):
    c = s(code)
    return _COUNTRY.get(c.upper(), c)

def ymd(v):
    v = s(v).split(".")[0].replace("-", "")
    if len(v) == 8 and v.isdigit():
        y, m, d = v[:4], v[4:6], v[6:8]
        if m == "00": return y                       # 연도만
        return f"{y}.{m}" if d == "00" else f"{y}.{m}.{d}"  # 일자 미상이면 연.월
    return s(v)

# 매칭 pid
jb = json.load(open("COMPA_통합best.json"))
pids = {str(t["과제고유번호"]) for e in jb.values() for t in e["top5"]}

p = pd.read_pickle(PF)
p["과제고유번호"] = p["과제고유번호"].astype(str)
sub = p[p["과제고유번호"].isin(pids)].copy()
del p

out = {}
for pid, g in sub.groupby("과제고유번호"):
    recs = g.to_dict("records")
    # 병합 키: 출원번호(있으면) > 등록번호 > 고유 인덱스
    groups = {}
    for i, r in enumerate(recs):
        appno, regno = s(r["출원번호"]), s(r["등록번호"])
        key = ("A", appno) if appno else (("R", regno) if regno else ("I", str(i)))
        groups.setdefault(key, []).append(r)
    pats = []
    for key, rows in groups.items():
        reg = next((r for r in rows if s(r["출원/등록 구분"]) == "등록"), None)
        app = next((r for r in rows if s(r["출원/등록 구분"]) == "출원"), None)
        base = reg or app or rows[0]
        pats.append({
            "특허명": s(base["발명의 명칭"]),
            "기관": s((app or base)["출원/등록 기관"]) or s(base["출원/등록 기관"]),
            "국가": country(base["출원/등록 국가코드"]),
            "출원번호": s(base["출원번호"]),
            "출원일": ymd((app or base)["출원일자"]) or ymd(base["출원일자"]),
            "등록번호": s(reg["등록번호"]) if reg else "",
            "등록일": ymd(reg["등록일자"]) if reg else "",
            "상태": "등록" if reg else "출원",
            "_yr": s(base["성과발생년도"]),
        })
    # 등록 우선 → 최신순
    pats.sort(key=lambda x: (0 if x["상태"] == "등록" else 1,
                             -(int(x["등록일"].replace(".", "")) if x["등록일"].replace(".", "").isdigit() else 0),
                             -(int(x["출원일"].replace(".", "")) if x["출원일"].replace(".", "").isdigit() else 0)))
    out[pid] = pats

json.dump(out, open(f"{SCRATCH}/pid_patents.json", "w"), ensure_ascii=False)
n_pat = sum(len(v) for v in out.values())
reg_n = sum(1 for v in out.values() for x in v if x["상태"] == "등록")
print(f"특허 보유 과제: {len(out)} | 병합 후 특허: {n_pat} (등록 {reg_n}/출원 {n_pat-reg_n})")
mx = max(out.items(), key=lambda kv: len(kv[1]))
print(f"최다 과제: {mx[0]} → {len(mx[1])}건")
print("샘플:", json.dumps(out[mx[0]][:2], ensure_ascii=False, indent=1))
