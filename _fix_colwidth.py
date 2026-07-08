# -*- coding: utf-8 -*-
"""'수요기술 목록' 표(번호/수요기술명/기업명)의 열 너비 조정:
 번호 좁게, 수요기술명 넓게, 기업명 유지. (총 너비 동일)"""
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC = sys.argv[1] if len(sys.argv) > 1 else "COMPA_최종Top5_보고서.docx"
DST = sys.argv[2] if len(sys.argv) > 2 else SRC

WIDTHS = [720, 5040, 2880]   # 번호 / 수요기술명 / 기업명  (합계 8640 = 기존과 동일)

d = Document(SRC)
n = 0
for tb in d.tables:
    hdr = tuple(c.text.strip() for c in tb.rows[0].cells)
    if hdr != ("번호", "수요기술명", "기업명"):
        continue
    tbl = tb._tbl
    # 고정 레이아웃
    tblPr = tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout"); tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW"); tblPr.append(tblW)
    tblW.set(qn("w:w"), str(sum(WIDTHS))); tblW.set(qn("w:type"), "dxa")
    # tblGrid
    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    for g, w in zip(cols, WIDTHS):
        g.set(qn("w:w"), str(w))
    # 각 셀 tcW
    for row in tb.rows:
        for cell, w in zip(row.cells, WIDTHS):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW"); tcPr.append(tcW)
            tcW.set(qn("w:w"), str(w)); tcW.set(qn("w:type"), "dxa")
    n += 1

d.save(DST)
print(f"열 너비 조정한 표: {n}개  (번호={WIDTHS[0]}, 수요기술명={WIDTHS[1]}, 기업명={WIDTHS[2]} dxa)")
print("saved:", DST)
