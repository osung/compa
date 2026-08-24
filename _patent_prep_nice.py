# -*- coding: utf-8 -*-
"""보고서용 특허 실적 준비 — nice_patent_overview_rnd_linked_abstract_promise.pkl 기준.

기존 _patent_prep.py(apollo df_pr_patent_260710_detail.pkl)를 대체한다. 산출 스키마는
동일해서 gen_report/gen_report_pdf 의 특허 실적 표가 그대로 동작한다.

  {과제고유번호: [{상태, 특허명, 기관, 국가, 출원일, 출원번호, 등록일, 등록번호}, ...]}

※ 유망성 점수는 절대로 싣지 않는다.
   원본에는 유망성점수·유망성_raw·모과제유망성·유망성점수_기본·법적상태_score·최신성_score·
   청구항_score·전략기술_보너스 같은 점수 컬럼이 들어 있는데, 이 값들은 로드 직후 버리고
   (DROP_COLS) 산출 직전에 다시 검사해 하나라도 남아 있으면 실행을 중단한다.
   보고서·LLM 프롬프트 어디에도 유망성 수치가 들어가지 않게 하기 위한 조치다.

사용: COMPA_SCRATCH=<scratch> COMPA_REPORT_JSON=MEDITEK_TOP10_보고서.json \
      python _patent_prep_nice.py [--src <pkl>] [--out <json>]
"""
import argparse
import json
import os
import re

import pandas as pd

SRC = "nice_patent_overview_rnd_linked_abstract_promise.pkl"
BEST = os.environ.get("COMPA_REPORT_JSON", "MEDITEK_TOP10_보고서.json")
SCRATCH = os.environ.get("COMPA_SCRATCH", ".")

# 점수·유망성 계열 — 로드 직후 폐기(보고서/프롬프트 유입 차단)
DROP_COLS = ["유망성점수", "유망성점수_기본", "유망성_raw", "모과제유망성",
             "법적상태_score", "최신성_score", "청구항_score", "전략기술_보너스",
             "전략기술_분야", "전략기술_12대", "전략기술_NEXT10", "전략기술_K문샷"]
# 실제로 쓰는 컬럼만 뽑아 쓴다(점수 컬럼 유입 차단 + itertuples 가 괄호 있는 컬럼명을
# 위치 이름으로 바꿔버리는 문제 회피 — '출원일자(YYYYMMDD)' 등은 식별자가 될 수 없다)
USE_COLS = {"과제번호": "과제번호", "특허명": "특허명", "출원기관명": "출원기관명",
            "특허출원번호": "출원번호", "출원일자(YYYYMMDD)": "출원일자",
            "특허등록번호": "등록번호", "특허등록일자(YYYYMMDD)": "등록일자",
            "특허등록상태명": "등록상태"}
# 실을 특허 상태 — 매칭 코퍼스 필터(match_meditek_supply.ALIVE_STATES)와 같은 규칙.
# 거절·취하·포기된 특허를 '출원'으로 표시하면 이전 협의가 가능한 것처럼 읽히고,
# 소멸된 특허를 '등록'으로 표시하면 살아 있는 권리처럼 읽힌다.
ALIVE_STATES = ("등록", "공개")
# 산출 레코드에 허용되는 키(이 외에는 싣지 않는다)
ALLOWED = {"상태", "특허명", "기관", "국가", "출원일", "출원번호", "등록일", "등록번호"}
_SCORE_HINT = re.compile(r"유망|score|점수|promise", re.I)


def s(v):
    v = str(v).strip()
    return "" if v in ("nan", "None", "NaT", "-") else v


def ymd(v):
    v = s(v).split(".")[0].replace("-", "")
    if len(v) == 8 and v.isdigit():
        y, m, d = v[:4], v[4:6], v[6:8]
        if m == "00":
            return y
        return f"{y}.{m}" if d == "00" else f"{y}.{m}.{d}"
    return s(v)


def pids_of(v):
    """과제번호 셀(리스트 또는 문자열) → 과제고유번호 목록."""
    if isinstance(v, (list, tuple, set)):
        return [s(x) for x in v if s(x)]
    return re.findall(r"\d{6,}", s(v))


def org(v):
    """출원기관명 'A ; B ; C' → 'A 외 2'(표 폭 유지). 단독이면 그대로."""
    parts = [x.strip() for x in s(v).split(";") if x.strip() and x.strip() != "-"]
    if not parts:
        return ""
    return parts[0] if len(parts) == 1 else f"{parts[0]} 외 {len(parts) - 1}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--best", default=BEST)
    ap.add_argument("--out", default=os.path.join(SCRATCH, "pid_patents.json"))
    ap.add_argument("--phase-map", default=os.environ.get(
        "COMPA_PHASE_MAP", "MEDITEK_260820_연차통합.json"),
        help="연차 통합 대응표 {대표 과제고유번호: [연차 과제고유번호…]}. "
             "다년차 과제의 특허를 대표 과제로 모아 싣는다.")
    a = ap.parse_args()

    jb = json.load(open(a.best, encoding="utf-8"))
    reps = {str(t["과제고유번호"]) for e in jb.values()
            for t in (e.get("top5") or e.get("top10") or [])}
    # 연차 통합: 대표 과제 하나에 전 연차의 특허를 모은다. 연차 번호 → 대표 번호 매핑.
    members = {}
    if a.phase_map and os.path.exists(a.phase_map):
        members = json.load(open(a.phase_map, encoding="utf-8"))
    rep_of = {}
    for rep in reps:
        for q in members.get(rep, [rep]):
            rep_of[str(q)] = rep
    pids = set(rep_of)
    n_multi = sum(1 for rep in reps if len(members.get(rep, [rep])) > 1)
    print(f"입력: {a.best} | 대표 과제 {len(reps)}건 · 연차 포함 과제번호 {len(pids)}건 "
          f"(다년차 {n_multi}건, 대응표 {a.phase_map if members else '없음'})")

    df = pd.read_pickle(a.src)
    dropped = [c for c in DROP_COLS if c in df.columns]
    df = df.drop(columns=dropped)                 # 유망성·점수 컬럼 즉시 폐기
    missing = [c for c in USE_COLS if c not in df.columns]
    if missing:
        raise SystemExit(f"[중단] 원본에 없는 컬럼: {missing}")
    df = df[list(USE_COLS)].rename(columns=USE_COLS)
    n0 = len(df)
    df = df[df["등록상태"].astype(str).str.strip().isin(ALIVE_STATES)]
    print(f"특허 상태 필터: {n0} → {len(df)}행 "
          f"(인정 {'/'.join(ALIVE_STATES)} · 배제 거절·취하·포기·소멸)")
    print(f"원본 {len(df)}행 · 폐기한 점수 컬럼 {len(dropped)}개: {dropped}")
    print(f"사용 컬럼: {list(df.columns)}")

    out = {}
    for r in df.itertuples(index=False):
        hit = {rep_of[p] for p in pids_of(r.과제번호) if p in rep_of}
        if not hit:
            continue
        regno, regdt = s(r.등록번호), ymd(r.등록일자)
        st = s(r.등록상태)
        rec = {
            # 상태명을 그대로 쓴다. '공개'는 출원 계류 중이라는 뜻이므로 '출원'으로 적는다.
            "상태": "등록" if st == "등록" else "출원",
            "특허명": s(r.특허명),
            "기관": org(r.출원기관명),
            "국가": "한국",                        # 국내 출원번호 체계(10…) 기준
            "출원일": ymd(r.출원일자),
            "출원번호": s(r.출원번호),
            "등록일": regdt,
            "등록번호": regno,
        }
        for p in hit:
            out.setdefault(p, []).append(rec)

    # 같은 과제 안에서 출원번호 중복 제거 → 등록 우선, 최신순
    def key(x):
        def num(v):
            v = v.replace(".", "")
            return int(v) if v.isdigit() else 0
        return (0 if x["상태"] == "등록" else 1, -num(x["등록일"]), -num(x["출원일"]))

    for p, lst in out.items():
        seen, uniq = set(), []
        for x in lst:
            k = x["출원번호"] or (x["특허명"], x["출원일"])
            if k in seen:
                continue
            seen.add(k)
            uniq.append(x)
        out[p] = sorted(uniq, key=key)

    # 안전장치: 허용 키 외 필드(=점수 유출) 검사
    bad = {k for lst in out.values() for x in lst for k in x if k not in ALLOWED}
    if bad:
        raise SystemExit(f"[중단] 허용되지 않은 필드가 산출에 포함됨: {sorted(bad)}")
    # 점수/유망성 흔적 검사 — 특허명은 원본 그대로 싣는 자유 텍스트라 검사에서 제외한다
    # (예: '딥러닝을 활용한 유전적 위험 점수 산출 장치 및 방법' 같은 실제 특허명이 걸린다).
    # 점수 컬럼 유입은 위의 ALLOWED 키 검사가 막으므로 이 검사는 그 외 필드만 본다.
    scan = json.dumps([{k: v for k, v in x.items() if k != "특허명"}
                       for lst in out.values() for x in lst], ensure_ascii=False)
    leak = _SCORE_HINT.findall(scan)
    if leak:
        raise SystemExit(f"[중단] 산출물에 점수/유망성 흔적: {set(leak)}")
    hint_names = [x["특허명"] for lst in out.values() for x in lst
                  if _SCORE_HINT.search(x["특허명"])]
    if hint_names:
        print(f"  (참고) '점수·유망' 문구가 든 특허명 {len(hint_names)}건 — 원본 특허명 그대로 유지: "
              + " / ".join(sorted(set(hint_names))[:3]))
    blob = json.dumps(out, ensure_ascii=False)

    with open(a.out, "w", encoding="utf-8") as f:
        f.write(blob)
    n = sum(len(v) for v in out.values())
    reg = sum(1 for v in out.values() for x in v if x["상태"] == "등록")
    print(f"특허 보유 과제 {len(out)}/{len(reps)} · 특허 {n}건(등록 {reg}/출원 {n - reg})")
    print(f"저장: {a.out}")
    if out:
        mx = max(out.items(), key=lambda kv: len(kv[1]))
        print(f"최다 과제 {mx[0]} → {len(mx[1])}건")
        print("샘플:", json.dumps(mx[1][:2], ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
