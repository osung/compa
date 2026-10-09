# -*- coding: utf-8 -*-
"""build_demand262.py 산출물(수요발굴_262_기업정보) 검증.

추출 때와 다른 방법(독립 정규식·좌표 계수·역조회)으로 원본 목록·PDF 조사서·KIPRIS/apollo 와
대조한다. 원본 자체의 이상으로 이미 '비고'에 표시된 건과 표기만 다른 같은 회사는
'확인된 예외'로 따로 보이고, 새로 생긴 문제만 NG 로 낸다(NG 가 있으면 종료코드 1).

사용: python verify_demand262.py [--out-prefix 수요발굴_262_기업정보]
"""
import argparse
import json
import os
import re
import sys
from difflib import SequenceMatcher

import pandas as pd
import pymupdf

from build_demand262 import APOLLO_CO, CORP, MANUAL, PDF, XLSX, corp_key, load_corp

ns = lambda s: re.sub(r"\s+", "", str(s))
_ng = 0


def report(name, rows, known=()):
    """rows 중 known(확인된 예외 키)에 든 것은 예외로, 나머지는 NG 로 낸다."""
    global _ng
    new = [r for r in rows if (r[0] if isinstance(r, tuple) else r) not in known]
    exc = [r for r in rows if (r[0] if isinstance(r, tuple) else r) in known]
    tag = "NG " if new else "OK "
    print(f"[{tag}] {name}: {len(new)}건 {new[:8] if new else ''}")
    if exc:
        print(f"      └ 확인된 예외 {len(exc)}건: {exc[:8]}")
    _ng += bool(new)


def seg(t, a, b):
    i = re.search(a, t)
    j = re.search(b, t[i.end():]) if i else None
    return t[i.end():i.end() + j.start()] if i and j else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out-prefix", default="수요발굴_262_기업정보")
    ap.add_argument("--xlsx", default=XLSX)
    ap.add_argument("--pdf", default=PDF)
    ap.add_argument("--corp", default=CORP)
    ap.add_argument("--apollo", default=APOLLO_CO)
    ap.add_argument("--manual", default=MANUAL)
    a = ap.parse_args()

    xl = pd.read_excel(a.out_prefix + ".xlsx", dtype=str).fillna("")
    pk = pd.read_pickle(a.out_prefix + ".pkl").astype(str).replace({"<NA>": "", "nan": ""})
    d = pymupdf.open(a.pdf)
    pages = {int(p): d[int(p) - 1].get_text() for p in xl["조사서 쪽"]}
    note = dict(zip(xl["기술번호"], xl["비고"]))

    # 0. xlsx == pkl
    report("xlsx·pkl 일치", [(i, c) for c in pk.columns for i in range(len(pk))
                             if ns(pk.at[i, c]) != ns(xl.at[i, c])])

    # 1. 원본 목록과 번호·기업명·기술번호·수요기술명 완전 일치
    src = pd.read_excel(a.xlsx, header=2)
    src = src[pd.to_numeric(src["번호"], errors="coerce").notna()]
    src["번호"] = src["번호"].astype(int).astype(str)
    m = src.merge(xl, on="번호", suffixes=("_s", ""))
    report("목록 행수", [] if len(m) == len(src) == len(xl) else [(len(src), len(m), len(xl))])
    report("목록 필드 일치", [(r["번호"], c) for _, r in m.iterrows()
                            for c in ["기업명", "기술번호", "수요기술명"]
                            if str(r[c + "_s"]).strip() != str(r[c]).strip()])
    report("기술번호 중복", xl["기술번호"][xl["기술번호"].duplicated()].tolist())

    # 2. 쪽 ↔ 기술번호
    report("쪽↔기술번호", [(r["기술번호"], r["조사서 쪽"]) for _, r in xl.iterrows()
                         if (re.findall(r"(?:BT|IT|NT|ET|융합)_\d+", pages[int(r["조사서 쪽"])][:300])
                             or [""])[0] != r["기술번호"]])
    body = {i + 1 for i, p in enumerate(d) if "수요기술 상세" in p.get_text()}
    report("조사서 쪽 누락/중복", sorted(body ^ set(pages)))

    # 3. 원문 보존: 공백 제거 후 PDF 구간과 문자 단위 동일
    checks = [("상세내용", r"수요기술\s*상세\s*\n?\s*내용", r"예상\s*적용", "수요기술 상세내용"),
              ("수요기술명(조사서)", r"기술 수요내용\s*\n\s*수요기술명", r"수요기술\s*분야", "수요기술명(조사서)"),
              ("적용제품", r"예상\s*적용\s*제품\s*및\s*서비스", r"수요유형", "예상 적용 제품 및 서비스")]
    for name, s0, s1, col in checks:
        report(f"{name} 원문 동일", [r["기술번호"] for _, r in xl.iterrows()
                                   if ns(seg(pages[int(r["조사서 쪽"])], s0, s1) or "\0") != ns(r[col])])
    report("내용/사양 ⊂ 상세내용", [r["기술번호"] for _, r in xl.iterrows() if r["수요기술 내용"] and not (
        ns(r["수요기술 내용"]) in ns(r["수요기술 상세내용"]) and ns(r["수요기술 사양"]) in ns(r["수요기술 상세내용"]))])

    # 4. 줄 잇기 결함·서식 잔재
    report("머리표 줄중간 결합", [(r["기술번호"], g.group(0)) for _, r in xl.iterrows()
                              for g in re.finditer(r"[^\s(][ \t]+[-○〇●▪][ \t]\S{0,8}", r["수요기술 상세내용"])])
    report("서식 잔재", [(r["기술번호"], c) for _, r in xl.iterrows()
                        for c in ["수요기술 상세내용", "예상 적용 제품 및 서비스", "수요기술명(조사서)",
                                  "국가과학기술표준분류(대)", "국가과학기술표준분류(중)"]
                        if re.search(r"[□■]|수요기술 분야|전략분야|대분류|중분류|수요유형|^내용", r[c])])

    # 5. 6T: ■ 재계산 + 기술번호 접두 (원본에서 접두와 다르게 체크한 건은 비고에 표시돼 있으면 예외)
    six = []
    for _, r in xl.iterrows():
        s = seg(pages[int(r["조사서 쪽"])], r"6T 기준", r"국가") or ""
        on = [x for x in ["IT", "BT", "NT", "ET", "융합"] if re.search(r"■\s*" + x, s)]
        if ", ".join(on) != r["수요기술 분야(6T)"] or r["기술번호"].split("_")[0] not in on:
            six.append((r["기술번호"], on))
    report("6T 일치(■·기술번호 접두)", six, known={k for k, v in note.items() if "6T 체크" in v})

    # 6. 전략분야 ■ 수 == 추출 항목 수
    report("전략분야 ■수=항목수", [(r["기술번호"], r["전략분야"]) for _, r in xl.iterrows()
                               if (seg(pages[int(r["조사서 쪽"])], r"전략분야", r"수요기술\s*상세") or "").count("■")
                               != len([x for x in r["전략분야"].split(", ") if x])])

    # 7. 체크 기호 수 == 선택 수
    ck, multi = [], []
    for _, r in xl.iterrows():
        p = d[int(r["조사서 쪽"]) - 1]
        y = p.search_for("예상 적용")[0].y1
        n = sum(1 for ch in "∨√✓✔" for q in p.search_for(ch) if q.y0 > y)
        n += sum(1 for w in p.get_text("words") if w[4] in ("V", "v") and w[1] > y)
        fixed = "(기술도입)" in p.get_text()                    # 구 서식: 수요유형 고정 표기
        k = sum(len([x for x in r[c].split(", ") if x])
                for c in ["기술도입 목적", "기술거래 희망 유형", "도입희망금액", "도입희망시기"])
        k += 0 if fixed else len(r["수요유형"].split(", "))
        if n != k:
            ck.append((r["기술번호"], n, k))
        multi += [(r["기술번호"], c) for c in ["도입희망금액", "도입희망시기", "수요유형"] if "," in r[c]]
    report("체크기호 수=선택 수", ck)
    report("단일선택 항목에 복수선택", multi)

    # 8. 수요기술명 목록≈조사서 (조사서 칸 복사 오류로 비고에 표시된 건은 예외)
    report("수요기술명 목록≈조사서(유사도<0.6)",
           [(r["기술번호"], round(SequenceMatcher(None, ns(r["수요기술명"]), ns(r["수요기술명(조사서)"])).ratio(), 2))
            for _, r in xl.iterrows()
            if SequenceMatcher(None, ns(r["수요기술명"]), ns(r["수요기술명(조사서)"])).ratio() < 0.6],
           known={k for k, v in note.items() if "조사서 수요기술명" in v})
    report("빈 필드", [(r["기술번호"], c) for _, r in xl.iterrows()
                      for c in ["수요기술 상세내용", "수요기술 분야(6T)", "국가과학기술표준분류(대)",
                                "국가과학기술표준분류(중)", "전략분야", "예상 적용 제품 및 서비스", "수요유형",
                                "기술도입 목적", "기술거래 희망 유형", "도입희망금액", "도입희망시기"] if not r[c].strip()])

    # 9. 사업자번호: 형식, KIPRIS·apollo 역조회
    corp = load_corp(a.corp)
    apb = set(pd.read_pickle(a.apollo)["사업자번호"])
    firms = xl.drop_duplicates("기업명")
    report("사업자번호 형식(000-00-00000)", [(r["기업명"], r["사업자번호"]) for _, r in firms.iterrows()
                                       if r["사업자번호"] and not re.fullmatch(r"\d{3}-\d{2}-\d{5}", r["사업자번호"])])
    back = []
    for _, r in firms[firms["사업자번호"] != ""].iterrows():
        s = corp[corp["사업자번호"] == r["사업자번호"]]
        if s.empty and r["사업자번호"].replace("-", "") not in apb:
            back.append((r["기업명"], "KIPRIS·apollo 어디에도 없음"))
        elif not s.empty and r["법인번호"] not in set(s["법인번호"]):
            back.append((r["기업명"], "법인번호 불일치"))
    report("사업자번호 역조회", back)
    # 한 사업자번호에 기업명이 여럿: 법인 형태 표기만 다른 같은 상호면 같은 회사(예외)
    g = firms[firms["사업자번호"] != ""].groupby("사업자번호")["기업명"].apply(list)
    multi = [(b, names) for b, names in g.items() if len(names) > 1]
    report("한 사업자번호에 여러 기업명", multi,
           known={b for b, names in multi if len({corp_key(n) for n in names}) == 1})
    report("같은 기업명 행 간 사업자번호 불일치",
           [n for n, k in xl.groupby("기업명")["사업자번호"].nunique().items() if k > 1])

    # 10. 근거 일관성·수동 선택 반영
    fill = ("일치", "단일", "적합도 1위", "수동 선택")
    report("사업자번호 유무↔근거 일관", [(r["기업명"], r["사업자번호 근거"]) for _, r in firms.iterrows()
                                   if bool(r["사업자번호"]) != any(t in r["사업자번호 근거"] for t in fill)
                                   and "사업자번호 없음" not in r["사업자번호 근거"]])
    report("검토 필요인데 후보 없음", [r["기업명"] for _, r in firms.iterrows()
                                 if "검토" in r["사업자번호 근거"] and not r["사업자번호 후보"]])
    report("자동 채움이 후보 1순위와 다름", [r["기업명"] for _, r in firms.iterrows()
                                     if r["사업자번호"] and r["사업자번호 후보"] and r["사업자번호 근거"] != "수동 선택"
                                     and not r["사업자번호 후보"].startswith(r["사업자번호"])])
    if os.path.exists(a.manual):
        sel = json.load(open(a.manual, encoding="utf-8"))
        u = firms.set_index("기업명")
        report("수동 선택값 반영", [(k, v) for k, v in sel.items()
                                 if k not in u.index or u.at[k, "사업자번호"] != v
                                 or not u.at[k, "사업자번호 근거"].startswith("수동")])

    print("\n검증 결과:", "NG 있음" if _ng else "모두 통과(확인된 예외 제외)")
    sys.exit(1 if _ng else 0)


if __name__ == "__main__":
    main()
