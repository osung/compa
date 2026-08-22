# -*- coding: utf-8 -*-
"""보고서용 논문 실적 준비 → pid_papers.json

  {과제고유번호: [{"논문명": …, "학술지명": …}, …]}

두 소스를 순서대로 시도한다.
  ① NTIS DB(vc_ntis_h_tb_0008) — 논문명 + 학술지명. 사내망에서만 닿는다.
  ② project_match_data_260612.pkl 의 '논문명_리스트' — 제목만(학술지명은 빈 값)

②로 만들면 학술지명 칸이 비므로 보고서는 논문명만 있는 표로 렌더링한다. 나중에 사내망에서
①로 다시 만들면 보고서 코드를 고치지 않고 학술지명 열이 채워진다.

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
ALLOWED = {"논문명", "학술지명"}
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


def target_pids(best):
    jb = json.load(open(best, encoding="utf-8"))
    return {str(t["과제고유번호"]) for e in jb.values()
            for t in (e.get("top5") or e.get("top10") or [])}


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
    a = ap.parse_args()

    pids = target_pids(a.best)
    print(f"입력: {a.best} | 대상 과제 {len(pids)}건")
    out = None if a.no_db else from_db(pids)
    src = "NTIS DB(논문명+학술지명)"
    if out is None:
        out = from_pkl(pids)
        src = "project_match_data(논문명만)"

    # 같은 과제 안에서 논문명 중복 제거(제목 기준), 제목 순 정렬
    for p, lst in out.items():
        seen, uniq = set(), []
        for x in lst:
            k = re.sub(r"\s+", "", x["논문명"]).lower()
            if k in seen:
                continue
            seen.add(k)
            uniq.append(x)
        out[p] = sorted(uniq, key=lambda x: x["논문명"])

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
