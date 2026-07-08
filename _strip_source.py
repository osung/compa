# -*- coding: utf-8 -*-
"""docx 보고서에서 '출처' 관련 내용 모두 제거:
 - Top 제목의 '(출처:...)' 접미사 제거
 - 분야 개요의 '· 출처 ...' 부분 제거
 - '· 후보 출처 태그 —' 범례 줄, '최종 추천 과제 출처 분포' 줄 문단째 삭제."""
import re, sys
from docx import Document

SRC = sys.argv[1] if len(sys.argv) > 1 else "COMPA_최종Top5_보고서.docx"
DST = sys.argv[2] if len(sys.argv) > 2 else SRC

d = Document(SRC)
n_head = n_gaeyo = n_del = 0
for p in list(d.paragraphs):
    t = p.text
    st = p.style.name if p.style is not None else ""
    ts = t.strip()
    if ts.startswith("· 후보 출처 태그") or ts.startswith("최종 추천 과제 출처 분포"):
        p._element.getparent().remove(p._element); n_del += 1
    elif st == "Heading 3" and "(출처" in t:
        for r in p.runs:
            if "(출처" in r.text:
                r.text = re.sub(r"\s*\(출처[:：][^)]*\)\s*$", "", r.text)
        n_head += 1
    elif "출처" in t and "개요" in t:
        for r in p.runs:
            if "출처" in r.text:
                r.text = re.sub(r"\s*·\s*출처.*$", "", r.text)
        n_gaeyo += 1

d.save(DST)
print(f"제목 접미사 제거:{n_head}  개요 출처 제거:{n_gaeyo}  문단 삭제:{n_del}")
print("saved:", DST)
