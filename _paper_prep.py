# -*- coding: utf-8 -*-
"""보고서용 논문 실적 준비 → pid_papers.json

  {과제고유번호: [{"논문명": …, "학술지명": …}, …]}

세 소스를 순서대로 시도한다.
  ① DB_paper_260824.pkl — NTIS 논문 성과 덤프(1,030,961행). 과제고유번호가 있고
     학술지명·주저자명·게재년도·DOI 까지 들어 있다. **기본 경로.**
  ② NTIS DB(vc_ntis_h_tb_0008) — 사내망에서만 닿는다.
  ③ project_match_data_260612.pkl 의 '논문명_리스트' — 제목만(나머지 칸은 빈 값)

③으로 만들면 학술지명 등이 비므로 보고서 표의 해당 칸이 비어 렌더링된다. 나중에 ①·②로
다시 만들면 보고서 코드를 고치지 않고 열이 채워진다.

※ 유망성 점수는 싣지 않는다(_patent_prep_nice.py 와 같은 방침). 논문 데이터에는 점수
  컬럼이 없지만, 산출 직전에 허용 키만 남는지 검사한다.

사용: COMPA_SCRATCH=<scratch> COMPA_REPORT_JSON=MEDITEK_260820_보고서.json \
      python _paper_prep.py [--out <json>] [--no-db]
"""
import argparse
import json
import os
import pickle
import re

BEST = os.environ.get("COMPA_REPORT_JSON", "MEDITEK_260820_보고서.json")
SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
# 보고서 표에 싣는 항목만 남긴다(그 외 필드는 산출 직전 검사에서 막는다).
# 중요한 것만 고른다 — 검증내역은 추천 과제 논문 368건이 전부 'SCIE' 라 정보가 없고,
# 권·호·시작페이지는 DOI 가 있으면 불필요하다.
ALLOWED = {"논문명", "학술지명", "주저자", "게재년도", "DOI"}
PAPER_PKL = os.environ.get("COMPA_PAPER_PKL", "DB_paper_260824.pkl")
_SCORE_HINT = re.compile(r"유망|score|점수|promise", re.I)

# NTIS DB (사내망) — apollo/create_paper_list_data.py 와 같은 접속 정보
DB_URL = os.environ.get("COMPA_NTIS_DB",
                        "mysql+pymysql://dev23_apollo_sdq@203.250.213.223:23306/apollo_test")
DB_PW = os.environ.get("COMPA_NTIS_DB_PW", "dev23_apollo_sdq!@")
PAPER_TABLE = "vc_ntis_h_tb_0008"
CONNECT_TIMEOUT = 15


def s(v):
    v = str(v).strip()
    return "" if v in ("nan", "None", "NaT", "-", "") else v


# 원본 주저자명에 Scopus 저자 ID 가 괄호로 붙어 있다 — 'Kim, Jeong Hun (57233549000)'.
# 보고서에는 사람 이름만 싣는다.
_AUTHOR_ID = re.compile(r"\s*\(\s*\d[\d\s]*\)\s*$")


def author(v):
    return _AUTHOR_ID.sub("", s(v)).strip()


def target_pids(best, phase_map=""):
    """추천 대표 과제 → (연차 포함 과제번호 집합, {연차번호: 대표번호}).

    다년차 통합 과제는 연차마다 논문이 따로 달려 있으므로 전 연차를 모아 대표 과제로
    합친다(특허 표와 같은 규칙 — _patent_prep_nice.py 참고).
    """
    jb = json.load(open(best, encoding="utf-8"))
    reps = {str(t["과제고유번호"]) for e in jb.values()
            for t in (e.get("top5") or e.get("top10") or [])}
    members = {}
    if phase_map and os.path.exists(phase_map):
        members = json.load(open(phase_map, encoding="utf-8"))
    rep_of = {}
    for rep in reps:
        for q in members.get(rep, [rep]):
            rep_of[str(q)] = rep
    return set(rep_of), rep_of


def from_paper_pkl(pids, src=PAPER_PKL):
    """DB_paper_260824.pkl → {과제고유번호: [{논문명·학술지명·주저자·게재년도·DOI}, …]}.

    itertuples 는 'DOI/url' 처럼 식별자가 될 수 없는 컬럼명을 위치 이름으로 바꿔버리므로
    컬럼을 먼저 표준 이름으로 바꿔서 쓴다(_patent_prep_nice.py 와 같은 방식).
    """
    if not os.path.exists(src):
        print(f"  · {src} 없음 → 다음 소스로")
        return None
    import pandas as pd
    use = {"과제고유번호": "pid", "논문명": "논문명", "학술지명": "학술지명",
           "주저자명": "주저자", "성과발생년도": "게재년도", "DOI/url": "DOI"}
    df = pd.read_pickle(src)
    miss = [c for c in use if c not in df.columns]
    if miss:
        print(f"  · {src} 에 없는 컬럼 {miss} → 다음 소스로")
        return None
    df = df[list(use)].rename(columns=use)
    df = df[df["pid"].astype(str).str.strip().isin(pids)]
    out = {}
    for r in df.itertuples(index=False):
        p, t = s(r.pid), s(r.논문명)
        if not (p and t):
            continue
        out.setdefault(p, []).append({
            "논문명": t, "학술지명": s(r.학술지명), "주저자": author(r.주저자),
            "게재년도": s(r.게재년도), "DOI": s(r.DOI),
        })
    print(f"  · {os.path.basename(src)}: 과제 {len(out)}개 · 논문 "
          f"{sum(map(len, out.values()))}건 (학술지명·주저자·게재년도·DOI 포함)")
    return out


def from_db(pids):
    """NTIS DB 에서 논문명·학술지명 조회. 닿지 않으면 None 을 돌려 폴백하게 한다."""
    try:
        from sqlalchemy import create_engine, inspect, text
    except ImportError:
        print("  · sqlalchemy 미설치 → DB 건너뜀")
        return None
    try:
        eng = create_engine(DB_URL, connect_args={"password": DB_PW,
                                                  "connect_timeout": CONNECT_TIMEOUT})
        insp = inspect(eng)
        # 컬럼명이 코드값이고 한글 이름은 주석에 있다(apollo 스크립트와 동일한 규약)
        col = {c.get("comment"): c["name"] for c in insp.get_columns(PAPER_TABLE)}
        need = ["과제고유번호", "논문명", "학술지명"]
        miss = [k for k in need if k not in col]
        if miss:
            print(f"  ! DB 테이블에 없는 컬럼: {miss} → 폴백")
            return None
        ids = ",".join(f"'{p}'" for p in sorted(pids))
        q = (f"select `{col['과제고유번호']}` as pid, `{col['논문명']}` as title, "
             f"`{col['학술지명']}` as journal from {PAPER_TABLE} "
             f"where `{col['과제고유번호']}` in ({ids})")
        out = {}
        with eng.connect() as conn:
            for pid, title, journal in conn.execute(text(q)):
                t = s(title)
                if len(t) < 10:                  # apollo 필터와 동일(제목 10자 미만 제외)
                    continue
                out.setdefault(str(pid), []).append({"논문명": t, "학술지명": s(journal)})
        print(f"  · DB 조회 성공: 과제 {len(out)}개 · 논문 {sum(map(len, out.values()))}건")
        return out
    except Exception as e:
        print(f"  ! DB 접속 실패({type(e).__name__}) → 폴백: {str(e)[:120]}")
        return None


def from_pkl(pids):
    """project_match_data 의 논문명_리스트(제목만)."""
    import compa_match as cm
    with open(cm.PROJECT_META, "rb") as f:
        pm = pickle.load(f)
    out = {}
    for p in pids:
        lst = (pm.get(p) or {}).get("논문명_리스트")
        if isinstance(lst, str):
            try:
                lst = eval(lst)                  # 원본이 문자열로 직렬화된 경우
            except Exception:
                lst = []
        for t in (lst or []):
            t = s(t)
            if t:
                out.setdefault(p, []).append({"논문명": t, "학술지명": ""})
    print(f"  · pkl 폴백: 과제 {len(out)}개 · 논문 {sum(map(len, out.values()))}건 "
          "(학술지명 없음)")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--best", default=BEST)
    ap.add_argument("--out", default=os.path.join(SCRATCH, "pid_papers.json"))
    ap.add_argument("--no-db", action="store_true", help="DB 조회를 건너뛰고 pkl 만 쓴다")
    ap.add_argument("--phase-map", default=os.environ.get(
        "COMPA_PHASE_MAP", "MEDITEK_260820_연차통합.json"),
        help="연차 통합 대응표 — 전 연차 논문을 대표 과제로 합친다")
    a = ap.parse_args()

    pids, rep_of = target_pids(a.best, a.phase_map)
    n_rep = len(set(rep_of.values()))
    print(f"입력: {a.best} | 대표 과제 {n_rep}건 · 연차 포함 과제번호 {len(pids)}건")
    out = from_paper_pkl(pids)
    src = f"{os.path.basename(PAPER_PKL)}(논문명·학술지명·주저자·게재년도·DOI)"
    if out is None and not a.no_db:
        out = from_db(pids)
        src = "NTIS DB(논문명+학술지명)"
    if out is None:
        out = from_pkl(pids)
        src = "project_match_data(논문명만)"

    if rep_of:                                  # 연차 → 대표 과제로 합친다
        merged = {}
        for q, lst in out.items():
            merged.setdefault(rep_of.get(str(q), str(q)), []).extend(lst)
        out = merged

    # 같은 과제 안에서 논문명 중복 제거(제목 기준), 제목 순 정렬
    for p, lst in out.items():
        seen, uniq = set(), []
        for x in lst:
            k = re.sub(r"\s+", "", x["논문명"]).lower()
            if k in seen:
                continue
            seen.add(k)
            uniq.append(x)
        # 최신 게재연도 우선, 같은 해는 제목순
        out[p] = sorted(uniq, key=lambda x: (-int(x.get("게재년도") or 0), x["논문명"]))

    bad = {k for lst in out.values() for x in lst for k in x if k not in ALLOWED}
    if bad:
        raise SystemExit(f"[중단] 허용되지 않은 필드: {sorted(bad)}")
    scan = json.dumps([{k: v for k, v in x.items() if k != "논문명"}
                       for lst in out.values() for x in lst], ensure_ascii=False)
    leak = _SCORE_HINT.findall(scan)
    if leak:
        raise SystemExit(f"[중단] 점수/유망성 흔적: {set(leak)}")

    with open(a.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    n = sum(len(v) for v in out.values())
    j = sum(1 for v in out.values() for x in v if x["학술지명"])
    print(f"소스: {src}")
    print(f"논문 보유 과제 {len(out)}/{len(pids)} · 논문 {n}건 (학술지명 있는 논문 {j}건)")
    print(f"저장: {a.out}")
    if out:
        mx = max(out.items(), key=lambda kv: len(kv[1]))
        print(f"최다 과제 {mx[0]} → {len(mx[1])}건")
        print("샘플:", json.dumps(mx[1][:2], ensure_ascii=False))


if __name__ == "__main__":
    main()
