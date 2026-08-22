# -*- coding: utf-8 -*-
"""MEDITEK 공급기관 매칭 보고서(docx·pdf) 품질 검증.

렌더링을 눈으로 확인하다 잡은 결함들을 코드로 고정해 둔 것이다. 회귀로 다시 새지 않게
매번 같은 항목을 기계로 검사한다. 실제로 이 검사들이 잡아낸 것:

  · 표시용 기술명이 '1' · '기술명' · '- SUS 코일…' 로 찍히던 문제
    (match_meditek_supply._first_line 수정 → LABEL_LIKE/최소 길이 검사로 고정)
  · 공급기관별 요약표의 특허 합계가 행 기준으로 중복 계산돼 본문 특허 실적과 어긋난 문제
    (과제 단위 집계로 수정 → 합계 일치 검사로 고정)
  · 제출 원문의 표기 오류가 그대로 실리던 문제
    (report_text_fix 적용 → 잔존 결함 패턴 검사로 고정)
  · 유망성 점수·적합도 수치·매칭 기준 표현이 새는지 (보고서 규칙)

사용: COMPA_SCRATCH=<scratch> python verify_report_supply.py
      [--json MEDITEK_260820_보고서.json] [--pdf …] [--docx …] [--pkl …]
종료코드 0=전항목 통과, 1=실패 있음.
"""
import argparse
import glob
import json
import os
import re
import sys

import pandas as pd

import report_text_fix as tf

SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
SRC_XLSX = "2026_MEDITEK_260820.xlsx"
TOP_KEY = "top5"

# 보고서에 새면 안 되는 표현(기존 MEDITEK 보고서 규칙)
FORBIDDEN = [
    (r"유망성\s*(점수|지수)|유망성\s*[:：]?\s*\d", "유망성 점수"),
    (r"적합도\s*[:：]?\s*\d", "적합도 수치"),
    (r"최종점수", "최종점수"),

    (r"기업DB", "질의 근거 라벨"),
    (r"공공\s*R&D", "공공 R&D 표기"),
    (r"\[사업내용\]|\[산업분류\]|\[기업 키워드\]", "합본 라벨 잔재"),
]
# 교정 후에도 남아 있으면 안 되는 표기 결함.
# 줄바꿈은 결함이 아니므로 공백류는 [ \t] 로만 본다 — \s 로 두면 필드 경계나 표 셀에서
# 줄이 바뀐 자리('…1' + 줄바꿈 + '개…')를 결함으로 잡는 오탐이 난다.
TEXT_DEFECTS = [
    (r"[가-힣][ \t]+(?:습?니다)(?=[\s.,)\]]|$)", "어미 분리"),
    (r"\d[ \t]+(?:년|건|명|개|종|배|차|회|개월|차원|%)", "숫자-단위 공백"),
    (r"\([ \t]|[ \t]\)", "괄호 안쪽 공백"),
    (r"[^\S\n]{2,}", "중복 공백"),
    (r"[ \t]+[,.]", "구두점 앞 공백"),
    (r",[ \t]*,", "중복 쉼표"),
    (r"[)\]][ \t]+(?:은|는|가|을|를|의|와|과|로|도)(?=[\s,.]|$)", "괄호 뒤 조사 분리"),
    (r"간겅|감연", "확인된 오탈자"),
]
SEP = "\n\uFFFF\n"   # 필드 이어 붙일 때 쓰는 경계 표시(패턴이 넘어가지 않게)
# 표시용 기술명이 이것뿐이면 도출 실패로 본다(제출 서식의 항목 라벨)
LABEL_LIKE = {"기술명", "수요기술명", "개요", "기업 개요", "기술 개요", "기업개요", "기술개요"}


def _pdf_text(path):
    from pypdf import PdfReader
    r = PdfReader(path)
    return len(r.pages), "\n".join(p.extract_text() or "" for p in r.pages)


def _docx_text(path):
    from docx import Document
    d = Document(path)
    body = [p.text for p in d.paragraphs]
    body += [c.text for t in d.tables for row in t.rows for c in row.cells]
    return len(d.tables), "\n".join(body)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", default="MEDITEK_260820_보고서.json")
    ap.add_argument("--pkl", default="MEDITEK_260820_매칭.pkl")
    ap.add_argument("--pdf", default="MEDITEK_공급기관_국가RnD_매칭보고서.pdf")
    ap.add_argument("--docx", default="")
    ap.add_argument("--patents", default=os.path.join(SCRATCH, "pid_patents.json"))
    ap.add_argument("--src", default=SRC_XLSX)
    ap.add_argument("--final", type=int, default=5)
    a = ap.parse_args()
    docx = a.docx or max(glob.glob("MEDITEK_공급기관_국가RnD_매칭보고서_v*.docx"),
                         key=lambda p: int(re.search(r"_v(\d+)\.docx$", p).group(1)),
                         default="")

    res, warn = [], []
    def chk(name, cond, detail=""):
        res.append((bool(cond), name, detail))

    J = json.load(open(a.json, encoding="utf-8"))
    PAT = json.load(open(a.patents, encoding="utf-8"))
    DF = pd.read_pickle(a.pkl)
    SUP = set(pd.read_excel(a.src, sheet_name="공급기관 리스트")["기관명"]
              .astype(str).str.strip())
    # 보고서가 싣는 것과 같은 교정본으로 비교한다(생성기도 같은 함수를 쓴다)
    tf.fix_report(J, PAT, top_key=TOP_KEY)
    tops = [t for v in J.values() for t in v[TOP_KEY]]
    pids = sorted({t["과제고유번호"] for t in tops})

    # ---- 규모·정합 ----
    chk("추천 규모", len(tops) == len(J) * a.final,
        f"기업 {len(J)} · 추천 {len(tops)} · 과제 {len(pids)}")
    chk("기업별 rank 1..n 연속",
        all(sorted(t["rank"] for t in v[TOP_KEY]) == list(range(1, a.final + 1))
            for v in J.values()))
    chk("기업별 과제 중복 없음",
        all(len({t["과제고유번호"] for t in v[TOP_KEY]}) == a.final for v in J.values()))
    pklmap = {(r["기업명"], int(r["rank"])): r for r in DF.to_dict("records")}
    mis = [(v["기업명"], t["rank"]) for v in J.values() for t in v[TOP_KEY]
           if (r := pklmap.get((v["기업명"], t["rank"]))) is None
           or str(r["과제고유번호"]) != t["과제고유번호"]
           or r["과제수행기관"] != t["수행기관"] or r["공급기관"] != t["공급기관"]]
    chk("보고서 JSON == 매칭 pkl", not mis, f"불일치 {mis[:3]}")

    # ---- 공급기관·특허 ----
    sups = {t.get("공급기관", "") for t in tops}
    chk("공급기관 전건 표기", all(t.get("공급기관") for t in tops), f"{len(sups)}곳")
    chk("공급기관이 리스트 소속", sups <= SUP, f"이탈 {sorted(sups - SUP)}")
    chk("모든 과제 특허 1건 이상", all(len(PAT.get(p, [])) >= 1 for p in pids),
        f"최소 {min(len(PAT.get(p, [])) for p in pids)}건 · "
        f"총 {sum(len(PAT.get(p, [])) for p in pids)}건")
    badp = [t["과제고유번호"] for t in tops
            if t["특허건수"] != len(PAT.get(t["과제고유번호"], []))]
    chk("표 특허건수 == 특허 실적 목록", not badp, f"불일치 {badp[:3]}")

    # ---- 표시용 기술명(과거 '1'·'기술명' 으로 찍힌 회귀 방지) ----
    bad = []
    for k, v in J.items():
        nm = (v.get("수요기술명") or v.get("기보유기술명") or v.get("기술분야") or "").strip()
        if not nm or len(nm) < 4 or nm in LABEL_LIKE or nm[0] in "-·•*1234567890":
            bad.append((k, nm))
    chk("표시용 기술명 정상", not bad, f"이상 {bad[:4]}")

    # ---- 근거문 ----
    chk("판단근거 전건", all(t["판단근거"].strip() for t in tops))
    # 기업 소개면에는 '수요기술/보유기술' 구분을 일부러 싣는다(요청). 다만 생성한 근거
    # 문장에는 매칭 기준 용어가 들어가면 안 된다 — 그 검사는 근거문 필드에만 적용한다.
    BAN = re.compile(r"수요기술|기보유기술|기술수요|확보하려는|수요\s?충족|보유\s?보강")
    bad_ban = [(k, t["rank"]) for k, v in J.items() for t in v[TOP_KEY]
               if BAN.search(t["판단근거"]) or BAN.search(t["추천근거_상세"])]
    chk("근거문에 매칭 기준 용어 없음", not bad_ban, f"위반 {bad_ban[:3]}")
    # 기업 소개면 구분 라벨이 실제로 실렸는지(요청 사항 확인)
    need = [v["기업명"] for v in J.values()
            if (v.get("수요기술 내용") or v.get("수요기술명"))]
    chk("수요기술 구분 라벨 대상 기업 존재", bool(need), f"{len(need)}개사")
    chk("상세근거 4섹션 전건",
        all(all(f"[{s}]" in t["추천근거_상세"]
                for s in ("연관성", "기술 적합성", "추천 과제의 우수성", "유사 사례 및 실적"))
            for t in tops))

    # ---- 표기 결함(교정 후 잔존) ----
    fields = [v.get(f, "") for v in J.values() for f in tf.COMPANY_FIELDS]
    fields += [t.get(f, "") for t in tops for f in tf.TOP_FIELDS]
    fields += [p.get(f, "") for lst in PAT.values() for p in lst for f in tf.PATENT_FIELDS]
    blob = SEP.join(x for x in fields if x)
    for pat, label in TEXT_DEFECTS:
        hit = [blob[max(0, x.start() - 14):x.end() + 10].replace("\n", "⏎")
               for x in re.finditer(pat, blob)]
        chk(f"표기 결함 없음: {label}", not hit, f"{len(hit)}건 {hit[:3]}")

    # ---- 산출물(pdf/docx) ----
    for label, path, reader in (("PDF", a.pdf, _pdf_text), ("DOCX", docx, _docx_text)):
        if not path or not os.path.exists(path):
            warn.append(f"{label} 파일 없음 → 건너뜀 ({path})")
            continue
        n, txt = reader(path)
        flat = txt.replace("\n", "").replace(" ", "")
        chk(f"{label} 생성", n > 0, f"{'쪽' if label == 'PDF' else '표'} {n}개")
        miss = [v["기업명"] for v in J.values() if v["기업명"].replace(" ", "") not in flat]
        chk(f"{label}: 기업 전수 수록", not miss, f"누락 {miss[:3]}" if miss else f"{len(J)}개사")
        mp = [p for p in pids if p not in flat]
        chk(f"{label}: 과제 전수 수록", not mp, f"누락 {mp[:3]}" if mp else f"{len(pids)}개")
        ms = [x for x in sups if x.replace(" ", "") not in flat]
        chk(f"{label}: 공급기관명 수록", not ms, f"누락 {ms}" if ms else f"{len(sups)}곳")
        for pat, lab in FORBIDDEN:
            m = re.findall(pat, txt)
            chk(f"{label}: {lab} 비노출", not m, f"{len(m)}건 {sorted(set(map(str, m)))[:3]}")
        tot = sum(len(PAT.get(p, [])) for p in pids)
        chk(f"{label}: 요약표 특허 합계 == 과제별 합계", f"{tot}건" in txt, f"{tot}건")
        chk(f"{label}: 깨진 글리프 없음", "□" not in txt, f"□ {txt.count('□')}개")
        chk(f"{label}: 수요기술·보유기술 구분 표기", "확보를 희망하는 기술" in txt
            and "이미 보유한 기술" in txt,
            f"수요 {txt.count('확보를 희망하는 기술')}회 · 보유 {txt.count('이미 보유한 기술')}회")
        # 렌더링본에서 뽑은 글은 표 셀·배지 같은 서로 다른 요소가 이어 붙는다
        # ('TOP 5' 배지 + 과제명 '차세대…' → '5 차'). 요소 인접만으로 생길 수 있는
        # 패턴(숫자-단위)은 여기서 보지 않는다 — 위의 필드 단위 검사가 이미 담당한다.
        for pat, lab in [TEXT_DEFECTS[0]] + TEXT_DEFECTS[5:8]:
            hit = [txt[max(0, x.start() - 16):x.end() + 12].replace("\n", "⏎")
                   for x in re.finditer(pat, txt)]
            if hit:
                warn.append(f"{label} 본문 '{lab}' {len(hit)}건 {hit[:2]} "
                            "— 교정 대상 외 필드(과제설명문 파생 등)일 수 있음")

    print("=" * 84)
    print("보고서 검증")
    print("=" * 84)
    for ok, name, detail in res:
        print(f"  [{'PASS' if ok else 'FAIL'}] {name:<44} {detail}")
    for w in warn:
        print(f"  [주의] {w}")
    nf = sum(1 for ok, _, _ in res if not ok)
    print(f"\n검증 {len(res)}항목 · 실패 {nf}건 · 주의 {len(warn)}건")
    return 0 if nf == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
