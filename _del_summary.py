# -*- coding: utf-8 -*-
"""'■ 전체 요약' 섹션(제목~다음 장 직전) 전체 삭제.
구역 나누기 sectPr는 전체 요약 앞 문단(구역1)에 있어 보존 → 제1장이 1페이지로 시작."""
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

SRC = sys.argv[1] if len(sys.argv) > 1 else "COMPA_최종Top5_보고서.docx"
DST = sys.argv[2] if len(sys.argv) > 2 else SRC

d = Document(SRC)
body = d.element.body
els = list(body)

start_i = end_i = None
for i, ch in enumerate(els):
    if ch.tag != qn("w:p"):
        continue
    p = Paragraph(ch, d)
    t = p.text.strip()
    st = p.style.name if p.style else ""
    if start_i is None and st == "Heading 1" and t.startswith("■ 전체 요약"):
        start_i = i
    elif start_i is not None and st == "Heading 1":   # 다음 Heading 1 = 다음 장
        end_i = i
        break

assert start_i is not None and end_i is not None, f"범위 못 찾음 start={start_i} end={end_i}"

# 안전장치: 삭제 대상에 sectPr(구역 나누기)가 없어야 함
for ch in els[start_i:end_i]:
    if ch.tag == qn("w:p"):
        pPr = Paragraph(ch, d)._p.pPr
        assert pPr is None or pPr.find(qn("w:sectPr")) is None, "삭제 범위에 구역 나누기 포함 — 중단"

removed = 0
for ch in els[start_i:end_i]:
    body.remove(ch); removed += 1

d.save(DST)
print(f"전체 요약 섹션 삭제: {removed}개 요소 (body {start_i}~{end_i-1})")
print("saved:", DST)
