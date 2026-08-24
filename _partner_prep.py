# -*- coding: utf-8 -*-
"""보고서용 공동참여기관 준비 → pid_partners.json

  {과제고유번호: [{"기관명": …, "참여형태": …}, …]}

소스: ../apollo/df_pr_outsource_260309.pkl
  컬럼 = 상위과제고유번호 · 공동위탁과제번호 · 참여형태 · 수행기관명 · 수행기관사업자등록번호
  참여형태 = 연구·기술개발 / 기술이전 및 사업화 / 인력양성 / 국제협약 / 기타

정제가 필요한 이유 — 원본에 같은 기관이 표기만 다르게 여러 번 들어 있다.
  '켐온' / '(주)켐온',  '마이크로핏' / '(주)마이크로핏',  '울산대학' / '울산대학교',
  같은 기관이 연차마다 재등재되어 'Dalhousie University' 가 두 번 나오는 식.
→ 정규화 키로 묶고, 표기는 **가장 긴 원문**을 남긴다((주) 가 붙은 정식 표기 선호).

주관기관(과제수행기관명)은 목록에서 뺀다 — 과제 정보표에 이미 별도 행으로 있다.

※ 커버리지는 낮다(추천 과제 기준 약 21%). 데이터가 있는 과제만 보고서에 행을 넣는다.
   시점도 260309 로 다른 소스(260813·260824)보다 오래됐다.

사용: COMPA_SCRATCH=<scratch> COMPA_REPORT_JSON=MEDITEK_260820_보고서.json \
      python _partner_prep.py [--src …] [--out …] [--phase-map …]
"""
import argparse
import json
import os
import re

SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
BEST = os.environ.get("COMPA_REPORT_JSON", "MEDITEK_260820_보고서.json")
SRC = os.environ.get("COMPA_PARTNER_PKL",
                     "/Users/osung/work/apollo/df_pr_outsource_260309.pkl")
ALLOWED = {"기관명", "참여형태"}
_SCORE_HINT = re.compile(r"유망|score|점수|promise", re.I)
# 표기만 다른 같은 기관을 묶기 위한 정규화
_DROP = re.compile(r"(산학협력단|학교법인|재단법인|재단|의료원|주식회사|\(주\)|㈜|\(사\)|"
                   r"\(재\)|\(유\)|\(합\))")
_UNIV = re.compile(r"대학교$")


def s(v):
    v = str(v).strip()
    return "" if v in ("nan", "None", "NaT", "-", "") else v


def key(name):
    """묶음 키 — 공백·법인 표기 제거 후 '대학교'→'대학' 으로 통일."""
    k = _DROP.sub("", re.sub(r"\s+", "", str(name)))
    return _UNIV.sub("대학", k)


def target_pids(best, phase_map=""):
    """추천 대표 과제 → (연차 포함 과제번호 집합, {연차번호: 대표번호})."""
    jb = json.load(open(best, encoding="utf-8"))
    reps = {str(t["과제고유번호"]) for e in jb.values()
            for t in (e.get("top5") or e.get("top10") or [])}
    own = {str(t["과제고유번호"]): s(t.get("수행기관", "")) for e in jb.values()
           for t in (e.get("top5") or e.get("top10") or [])}
    members = {}
    if phase_map and os.path.exists(phase_map):
        members = json.load(open(phase_map, encoding="utf-8"))
    rep_of = {}
    for rep in reps:
        for q in members.get(rep, [rep]):
            rep_of[str(q)] = rep
    return rep_of, own


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--best", default=BEST)
    ap.add_argument("--out", default=os.path.join(SCRATCH, "pid_partners.json"))
    ap.add_argument("--phase-map", default=os.environ.get(
        "COMPA_PHASE_MAP", "MEDITEK_260820_연차통합.json"))
    a = ap.parse_args()

    if not os.path.exists(a.src):
        raise SystemExit(f"[중단] 소스 없음: {a.src}")
    import pandas as pd
    rep_of, own = target_pids(a.best, a.phase_map)
    reps = set(rep_of.values())
    df = pd.read_pickle(a.src)
    col = "상위과제고유번호" if "상위과제고유번호" in df.columns else "과제고유번호"
    nmc = next(c for c in df.columns if "기관명" in c)
    print(f"입력: {os.path.basename(a.src)} {df.shape} | 대표 과제 {len(reps)}건 · "
          f"연차 포함 {len(rep_of)}건")
    df = df[df[col].astype(str).str.strip().isin(rep_of)]

    raw, out = 0, {}
    for r in df.itertuples(index=False):
        rep = rep_of[str(getattr(r, col)).strip()]
        nm = s(getattr(r, nmc))
        if not nm:
            continue
        raw += 1
        form = s(getattr(r, "참여형태", ""))
        d = out.setdefault(rep, {})
        k = key(nm)
        if k in d:
            # 표기는 가장 긴 원문을 남기고, 참여형태는 합집합으로 모은다
            if len(nm) > len(d[k]["기관명"]):
                d[k]["기관명"] = nm
            if form:
                d[k]["_forms"].add(form)
        else:
            d[k] = {"기관명": nm, "_forms": {form} if form else set()}

    # 주관기관 제거 + 정렬 + 참여형태 확정
    final, dropped_own = {}, 0
    for rep, d in out.items():
        mine = key(own.get(rep, ""))
        lst = []
        for k, v in d.items():
            if mine and k == mine:
                dropped_own += 1
                continue
            lst.append({"기관명": v["기관명"],
                        "참여형태": " · ".join(sorted(v["_forms"]))})
        if lst:
            final[rep] = sorted(lst, key=lambda x: x["기관명"])

    bad = {k for lst in final.values() for x in lst for k in x if k not in ALLOWED}
    if bad:
        raise SystemExit(f"[중단] 허용되지 않은 필드: {sorted(bad)}")
    leak = _SCORE_HINT.findall(json.dumps(final, ensure_ascii=False))
    if leak:
        raise SystemExit(f"[중단] 점수/유망성 흔적: {set(leak)}")

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(final, f, ensure_ascii=False)
    n = sum(len(v) for v in final.values())
    print(f"원본 행 {raw} → 정제 {n}건 · 과제 {len(final)}/{len(reps)} "
          f"({len(final) / max(len(reps), 1) * 100:.0f}%) · 주관기관 제외 {dropped_own}건")
    print(f"저장: {a.out}")
    if final:
        mx = max(final.items(), key=lambda kv: len(kv[1]))
        print(f"최다 과제 {mx[0]} → {len(mx[1])}개")
        for x in mx[1]:
            print(f"   · {x['기관명']}  [{x['참여형태']}]")


if __name__ == "__main__":
    main()
