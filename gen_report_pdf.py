# -*- coding: utf-8 -*-
"""COMPA 최종 매칭 보고서 PDF 직접 생성(reportlab) — docx(gen_report.py) 출판물 디자인 재현.
Noto Sans KR(표제·표) + Noto Serif KR(본문 프로즈), 러닝헤더+APOLLO 로고, 표지 메타카드,
섹션바(▍), 장 영문캡션, TOP 배지, 4열 과제 정보표(연구책임자·국가연구자번호 포함).
Word/LibreOffice 불필요. 입력: COMPA_통합best.json + scratchpad/{pid_fields,demand_field}.json"""
import json, os, re
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, PageBreak, KeepTogether, NextPageTemplate)
from reportlab.lib.styles import ParagraphStyle

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.environ.get("COMPA_SCRATCH",
    "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
OUT = os.environ.get("COMPA_PDF_OUT", os.path.join(HERE, "COMPA_필터전체_보고서.pdf"))
FDIR = os.path.join(SCRATCH, "fonts")
LOGO = os.path.join(SCRATCH, "apollo_top.png")

# ---- 폰트 (Noto Sans/Serif KR, Regular+Bold) ----
for nm, fn in [("Sans", "NotoSansKR-Regular.ttf"), ("Sans-B", "NotoSansKR-Bold.ttf"),
               ("Serif", "NotoSerifKR-Regular.ttf"), ("Serif-B", "NotoSerifKR-Bold.ttf")]:
    pdfmetrics.registerFont(TTFont(nm, os.path.join(FDIR, fn)))

def F(family, bold):
    return ("Sans-B" if bold else "Sans") if family == "Sans" else ("Serif-B" if bold else "Serif")

# ---- 팔레트 (docx와 동일) ----
INK = colors.HexColor("#1B2430"); NAVY = colors.HexColor("#14315C"); BLUE = colors.HexColor("#2C5FA0")
ACCENT = colors.HexColor("#0E7C86"); MUTED = colors.HexColor("#6B7683"); HAIR = colors.HexColor("#C9D2DE")
HEADBG = colors.HexColor("#14315C"); HEADFG = colors.white; LABELBG = colors.HexColor("#EAF0F7")
ZEBRA = colors.HexColor("#F5F8FC"); DISCBG = colors.HexColor("#FBEEED"); DISCBD = colors.HexColor("#C0392B")

FIELD_ORDER = ["BT", "IT", "NT", "ET", "융합"]
FIELD_TITLE = {"BT": "바이오기술 (BT) 분야", "IT": "정보기술 (IT) 분야", "NT": "나노기술 (NT) 분야",
               "ET": "환경기술 (ET) 분야", "융합": "융합기술 분야"}
FIELD_EN = {"BT": "BIOTECHNOLOGY", "IT": "INFORMATION TECHNOLOGY", "NT": "NANOTECHNOLOGY",
            "ET": "ENVIRONMENTAL TECHNOLOGY", "융합": "CONVERGENCE TECHNOLOGY"}
PUBLISH_DATE = "2026. 7. 12."
DISCLAIMER = ("본 보고서는 APOLLO 인공지능을 이용해서 생성한 보고서로 사실과 다르거나 오류가 있을 수 있습니다. "
              "참고용으로만 활용하시고, 정확한 정보는 관련 자료를 통해 확인하시기 바랍니다. 본 보고서는 AI 생성 "
              "내용의 정확성을 보증하지 않으며, 이를 근거로 한 판단 의사결정의 책임은 이용자에게 있습니다.")

def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>"))

def P(txt, size=9, color=INK, align=TA_LEFT, leading=None, bold=False, family="Sans", space=0, indent=0):
    st = ParagraphStyle("s", fontName=F(family, bold), fontSize=size, textColor=color, alignment=align,
                        leading=leading or size * 1.35, spaceAfter=space, leftIndent=indent)
    return Paragraph(esc(txt), st)

# ---- 데이터 ----
demands = json.load(open(os.environ.get("COMPA_REPORT_JSON", os.path.join(HERE, "COMPA_통합best.json")), encoding="utf-8"))
pidf = json.load(open(os.path.join(SCRATCH, "pid_fields.json"), encoding="utf-8"))
field6t = json.load(open(os.path.join(SCRATCH, "demand_field.json"), encoding="utf-8"))
_pp = os.path.join(SCRATCH, "pid_patents.json")
patents = json.load(open(_pp, encoding="utf-8")) if os.path.exists(_pp) else {}

def fmt_period(desc):
    m = re.search(r'(\d{4})년\s*\d{1,2}월\s*\d{1,2}일에 시작.*?(\d{4})년\s*\d{1,2}월\s*\d{1,2}일에 종료', desc or "")
    if m: return m.group(1) + "년" if m.group(1) == m.group(2) else f"{m.group(1)}년 ~ {m.group(2)}년"
    m2 = re.search(r'(\d{4})년', desc or ""); return (m2.group(1) + "년") if m2 else ""

def year_cell(desc):
    m = re.search(r'(\d{4})년\s*\d{1,2}월\s*\d{1,2}일에 시작.*?(\d{4})년\s*\d{1,2}월\s*\d{1,2}일에 종료', desc or "")
    if m: return m.group(1) if m.group(1) == m.group(2) else f"{m.group(1)} ~\n{m.group(2)}"
    m2 = re.search(r'(\d{4})년', desc or ""); return m2.group(1) if m2 else ""

def extract_class(desc):
    for p in [r'[,，]\s*([가-힣][가-힣·.\s]*?)\s*분야에\s*(?:속|해당)',
              r'([가-힣][가-힣·.\s]{1,18}?)\s*분야에\s*(?:속|해당)']:
        m = re.search(p, desc or "")
        if m:
            v = re.sub(r'\s*\.\s*', '·', m.group(1).strip().rstrip('.')).strip('· ')
            if 1 < len(v) <= 25 and not any(x in v for x in ("과제", "수행", "연구비")): return v
    return ""

def split_sections(detail):
    parts = re.split(r'(\[[^\]]+\])', detail or ""); out = []
    for i in range(1, len(parts), 2):
        out.append((parts[i].strip("[]"), parts[i + 1].strip() if i + 1 < len(parts) else ""))
    return out

# ---- 페이지 ----
PAGE_W, PAGE_H = A4
LM, RM, TM, BM = 18 * mm, 18 * mm, 22 * mm, 16 * mm
CW = PAGE_W - LM - RM
LOGO_W = 29 * mm; LOGO_H = LOGO_W * 276 / 1295

def draw_logo(c, y):
    try: c.drawImage(LOGO, PAGE_W - RM - LOGO_W, y, width=LOGO_W, height=LOGO_H, mask="auto")
    except Exception: pass

class Doc(BaseDocTemplate):
    def __init__(self, fn):
        super().__init__(fn, pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=TM, bottomMargin=BM)
        fr = Frame(LM, BM, CW, PAGE_H - TM - BM, id="n", topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate(id="cover", frames=[fr], onPage=self._cover),
                               PageTemplate(id="body", frames=[fr], onPage=self._body)])
    def _cover(self, c, d):
        draw_logo(c, PAGE_H - TM + 4)
    def _body(self, c, d):
        c.saveState()
        c.setFont("Sans", 8); c.setFillColor(MUTED)
        c.drawString(LM, PAGE_H - TM + 7, "COMPA  기술수요–공공 R&D 매칭 보고서")
        draw_logo(c, PAGE_H - TM + 3)
        c.setStrokeColor(HAIR); c.setLineWidth(0.6); c.line(LM, PAGE_H - TM, PAGE_W - RM, PAGE_H - TM)
        n = d.page - COVER_PAGES
        if n >= 1:
            c.setFont("Sans", 9); c.setFillColor(MUTED)
            c.drawCentredString(PAGE_W / 2, BM - 9, f"— {n} —")
        c.restoreState()

COVER_PAGES = 2
story = []

def base_grid(extra=None, fontsize=9):
    s = [("GRID", (0, 0), (-1, -1), 0.5, HAIR), ("FONT", (0, 0), (-1, -1), "Sans", fontsize),
         ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3.2),
         ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2), ("LEFTPADDING", (0, 0), (-1, -1), 5),
         ("RIGHTPADDING", (0, 0), (-1, -1), 5)]
    return s + (extra or [])

def mktable(data, widths, style):
    t = Table(data, colWidths=widths); t.setStyle(TableStyle(style)); return t

def section_label(text, before=10, after=5, size=13.5):
    st = ParagraphStyle("sl", fontName="Sans-B", fontSize=size, textColor=NAVY,
                        leading=size * 1.3, spaceBefore=before, spaceAfter=after)
    return Paragraph(f'<font color="#0E7C86">▍</font> {esc(text)}', st)

def cover():
    story.append(Spacer(1, 34 * mm))
    story.append(P("TECHNOLOGY  DEMAND  ×  PUBLIC  R&D  MATCHING", 10.5, ACCENT, TA_CENTER, bold=True, space=10))
    story.append(P("COMPA 매칭데이", 27, NAVY, TA_CENTER, bold=True, space=3))
    story.append(P("기술수요조사 최종 매칭 보고서", 27, NAVY, TA_CENTER, bold=True, space=10))
    story.append(mktable([[""]], [70 * mm], [("LINEBELOW", (0, 0), (-1, -1), 1.5, NAVY)]))
    story.append(Spacer(1, 5 * mm))
    story.append(P("기업 진성수요 × 공공 R&D 과제, 의미 기반 매칭 결과", 12, MUTED, TA_CENTER, space=22))
    n_dem = len(demands); n_rec = sum(len(v["top5"]) for v in demands.values())
    meta = mktable([[P(f"{n_dem}건", 20, NAVY, TA_CENTER, bold=True), P(f"{n_rec}건", 20, NAVY, TA_CENTER, bold=True)],
                    [P("대상 수요기술", 9.5, MUTED, TA_CENTER), P("추천 과제", 9.5, MUTED, TA_CENTER)]],
                   [45 * mm, 45 * mm],
                   [("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)])
    meta.hAlign = "CENTER"; story.append(meta)
    story.append(Spacer(1, 12 * mm))
    story.append(P(f"발행일  {PUBLISH_DATE}       생성  APOLLO AI 매칭 엔진", 10, MUTED, TA_CENTER, space=16))
    disc = mktable([[Paragraph(f'<font name="Sans-B" color="#A93226" size="10.5">※  유의사항</font><br/>'
                               f'<font name="Sans" color="#7B241C" size="9.5">{esc(DISCLAIMER)}</font>',
                               ParagraphStyle("d", leading=14))]], [CW],
                   [("BOX", (0, 0), (-1, -1), 0.8, DISCBD), ("BACKGROUND", (0, 0), (-1, -1), DISCBG),
                    ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                    ("LEFTPADDING", (0, 0), (-1, -1), 11), ("RIGHTPADDING", (0, 0), (-1, -1), 11)])
    story.append(disc)
    story.append(NextPageTemplate("body"))
    story.append(PageBreak())

def intro_toc(by_field):
    n_dem = len(demands); n_rec = sum(len(v["top5"]) for v in demands.values())
    n_proj = len({t["과제고유번호"] for v in demands.values() for t in v["top5"]})
    n_fields = len([f for f in FIELD_ORDER if by_field[f]])
    story.append(section_label("개요", before=2))
    story.append(P(f"본 보고서는 매칭데이 기술수요조사를 통해 취합된 기업 진성수요 {n_dem}건을 대상으로, 공공 R&D 과제 "
                   f"데이터베이스와의 의미 기반 매칭을 수행한 결과를 정리한 것이다. 각 수요기술에 대해 적합도가 높은 추천 "
                   f"과제 상위 5건(총 {n_rec}건, 중복 제외 {n_proj}개 과제)을 선정하고, 매칭 근거와 상세 추천 근거를 함께 "
                   f"제시하였다. 추천 대상은 최근 5년 이내(제출년도 2020년 이후) 과제이며, 연구수행주체가 대학·출연연구소·"
                   f"국공립연구소·정부부처인 과제로 한정하였다. 수요는 6T 기술 분류에 따라 {n_fields}개 분야로 구분하여 "
                   f"수록하였다.", 10.5, INK, TA_JUSTIFY, leading=17, family="Serif", space=8))
    story.append(P("각 수요기술은 다음 순서로 구성된다.", 10.5, INK, family="Serif", space=3))
    for ln in ["수요 정보 — 기업명 · 수요기술 내용 · 수요기술 사양",
               "최종 추천 과제 Top 5 — 순위 · 과제명 · 수행기관 · 과제수행년도 · 매칭 근거",
               "추천 과제별 상세 정보표 및 상세 매칭 근거 — 연관성 · 수요기술 사양 적합성 · 추천 과제의 우수성 · 유사 사례 및 실적"]:
        story.append(Paragraph(f'<font name="Sans-B" color="#0E7C86">· </font>'
                               f'<font name="Serif" color="#1B2430">{esc(ln)}</font>',
                               ParagraphStyle("b", fontSize=10, leading=15, leftIndent=14, spaceAfter=3)))
    story.append(section_label("목차", before=16))
    rows = []
    for i, f in enumerate([f for f in FIELD_ORDER if by_field[f]], 1):
        ks = by_field[f]; rng = f"수요 {ks[0]}–{ks[-1]}" if len(ks) > 1 else f"수요 {ks[0]}"
        rows.append([P(f"제 {i} 장", 10.5, ACCENT, bold=True), P(FIELD_TITLE[f], 11.5, NAVY, bold=True),
                     P(f"{rng} · {len(ks)}건", 9.5, MUTED, TA_RIGHT)])
    story.append(mktable(rows, [20 * mm, CW - 62 * mm, 42 * mm],
                 [("LINEBELOW", (0, 0), (-1, -1), 0.4, HAIR), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                  ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    story.append(PageBreak())

def chapter(no, f, ks):
    story.append(P(f"제{no}장  {FIELD_TITLE[f]}", 21, NAVY, bold=True, space=1))
    story.append(mktable([[P(FIELD_EN[f], 9, ACCENT, bold=True)]], [CW],
                 [("LINEBELOW", (0, 0), (-1, -1), 1.2, NAVY), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                  ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(Spacer(1, 3))
    story.append(P(f"수요기술 {len(ks)}건  ·  수요 {ks[0]}–{ks[-1]}" if len(ks) > 1 else f"수요기술 {len(ks)}건",
                   10, MUTED, space=10))
    story.append(section_label("수요기술 목록", before=2, after=6))
    rows = [[P(x, 9.5, HEADFG, TA_CENTER, bold=True) for x in ("번호", "수요기술명", "기업명")]]
    for k in ks:
        rows.append([P(k, 9.5, NAVY, TA_CENTER, bold=True), P(demands[k]["수요기술명"], 9.5, leading=12),
                     P(demands[k]["기업명"], 9.5, INK, TA_CENTER)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)])
    for i in range(2, len(rows), 2): st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    story.append(mktable(rows, [14 * mm, CW - 54 * mm, 40 * mm], st))
    story.append(PageBreak())
    for k in ks:
        demand_block(k, demands[k])

def demand_block(k, dm):
    badge = (f'<font name="Sans-B" color="#FFFFFF" backColor="#0E7C86"> 수요 {esc(k)} </font>'
             f'  <font name="Sans-B" color="#14315C" size="14">{esc(dm["수요기술명"])}</font>')
    story.append(Paragraph(badge, ParagraphStyle("h2", fontName="Sans-B", fontSize=14, leading=20, spaceAfter=5)))
    story.append(mktable([[""]], [CW], [("LINEBELOW", (0, 0), (-1, -1), 0.6, HAIR)]))
    story.append(Spacer(1, 4))
    rows = [[P("기업명", 9.5, NAVY, TA_CENTER, bold=True), P(dm.get("기업명", ""), 9)]]
    if dm.get("수요기술 내용", "").strip():
        rows.append([P("수요기술 내용", 9.5, NAVY, TA_CENTER, bold=True), P(dm["수요기술 내용"].strip(), 9, INK, TA_JUSTIFY, family="Serif", leading=13)])
    if dm.get("수요기술 사양", "").strip():
        rows.append([P("수요기술 사양", 9.5, NAVY, TA_CENTER, bold=True), P(dm["수요기술 사양"].strip(), 9, INK, TA_JUSTIFY, family="Serif", leading=13)])
    st = base_grid()
    for i in range(len(rows)): st.append(("BACKGROUND", (0, i), (0, i), LABELBG))
    story.append(mktable(rows, [28 * mm, CW - 28 * mm], st))
    story.append(Spacer(1, 6))
    story.append(section_label("최종 추천 과제  Top 5", before=4, after=5, size=12))
    rows = [[P(x, 9, HEADFG, TA_CENTER, bold=True) for x in ("순위", "과제명", "수행기관", "수행년도", "매칭 근거")]]
    for tp in dm["top5"]:
        rows.append([P(str(tp["rank"]), 10.5, NAVY, TA_CENTER, bold=True), P(tp["과제명"], 8.6, leading=11),
                     P(tp.get("수행기관", ""), 8.6, INK, TA_CENTER), P(year_cell(tp.get("과제설명문", "")), 8.6, INK, TA_CENTER),
                     P(tp.get("판단근거", ""), 8.6, INK, TA_JUSTIFY, family="Serif", leading=11)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)])
    for i in range(2, len(rows), 2): st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    story.append(mktable(rows, [12 * mm, 58 * mm, 26 * mm, 18 * mm, CW - 114 * mm], st))
    story.append(PageBreak())
    for tp in dm["top5"]:
        top_detail(tp)

def top_detail(tp):
    pid = str(tp["과제고유번호"]); ex = pidf.get(pid, {})
    title = (f'<font name="Sans-B" color="#FFFFFF" backColor="#14315C"> TOP {tp["rank"]} </font>'
             f'  <font name="Sans-B" color="#1B2430" size="12.5">{esc(tp["과제명"])}</font>')
    block = [Paragraph(title, ParagraphStyle("h3", fontName="Sans-B", fontSize=12.5, leading=18, spaceAfter=5)),
             mktable([[""]], [CW], [("LINEBELOW", (0, 0), (-1, -1), 0.6, HAIR)]), Spacer(1, 5)]
    info = [("과제고유번호", pid)]
    if fmt_period(tp.get("과제설명문", "")): info.append(("과제수행기간", fmt_period(tp.get("과제설명문", ""))))
    if extract_class(tp.get("과제설명문", "")): info.append(("과학기술표준분류(중)", extract_class(tp.get("과제설명문", ""))))
    if ex.get("연구개발단계"): info.append(("연구개발단계", ex["연구개발단계"]))
    info.append(("과제수행기관", tp.get("수행기관", "")))
    if ex.get("연구책임자명"): info.append(("연구책임자", ex["연구책임자명"]))
    if ex.get("국가연구자번호"): info.append(("국가연구자번호", ex["국가연구자번호"]))
    if ex.get("연구수행주체"): info.append(("연구수행주체", ex["연구수행주체"]))
    LW = 33 * mm; VW = (CW - 2 * LW) / 2
    rows = []
    for i in range(0, len(info), 2):
        l1, v1 = info[i]; cell = [P(l1, 9, NAVY, bold=True), P(str(v1), 9)]
        if i + 1 < len(info):
            l2, v2 = info[i + 1]; cell += [P(l2, 9, NAVY, bold=True), P(str(v2), 9)]
        else:
            cell += [P("", 9), P("", 9)]
        rows.append(cell)
    st = base_grid([("BACKGROUND", (0, 0), (0, -1), LABELBG), ("BACKGROUND", (2, 0), (2, -1), LABELBG)])
    if len(info) % 2 == 1:
        st.append(("BACKGROUND", (2, -1), (3, -1), colors.white))
    block.append(mktable(rows, [LW, VW, LW, VW], st))
    block.append(Spacer(1, 5))
    block.append(Paragraph(f'<font name="Sans-B" color="#FFFFFF" backColor="#2C5FA0"> 적합성 판단 </font>'
                           f'  <font name="Sans-B" color="#1B2430" size="9.5">{esc(tp.get("판단근거",""))}</font>',
                           ParagraphStyle("fit", fontSize=9.5, leading=14, spaceAfter=4)))
    block.append(section_label("상세 매칭 근거", before=4, after=3, size=10.5))
    for tt, body in split_sections(tp.get("추천근거_상세", "")):
        block.append(Paragraph(f'<font name="Sans-B" color="#2C5FA0">[{esc(tt)}]</font>  '
                               f'<font name="Serif" color="#1B2430">{esc(body)}</font>',
                               ParagraphStyle("sec", fontSize=9, leading=13.5, alignment=TA_JUSTIFY, spaceAfter=4)))
    story.append(KeepTogether(block))
    # ---- 특허 실적(있는 경우): 등록 우선, 출원정보 병기. 다년도 전 연도 포함 ----
    pats = patents.get(pid, [])
    if pats:
        story.append(section_label(f"특허 실적  ({len(pats)}건)", before=6, after=3, size=10.5))
        head = [P(x, 8, HEADFG, TA_CENTER, bold=True) for x in
                ("구분", "특허명", "출원·등록기관", "국가", "출원일", "출원번호", "등록일", "등록번호")]
        rows = [head]
        for pt in pats:
            reg = pt["상태"] == "등록"
            rows.append([
                P(pt["상태"], 8, (NAVY if reg else MUTED), TA_CENTER, bold=reg),
                P(pt["특허명"], 8, INK, leading=10),
                P(pt["기관"], 8, INK, TA_CENTER, leading=10),
                P(pt["국가"], 8, INK, TA_CENTER),
                P(pt["출원일"], 8, INK, TA_CENTER),
                P(pt["출원번호"], 8, INK, TA_CENTER),
                P(pt["등록일"], 8, INK, TA_CENTER),
                P(pt["등록번호"], 8, INK, TA_CENTER)])
        st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)], fontsize=8)
        for i in range(2, len(rows), 2): st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
        story.append(mktable(rows, [9 * mm, 44 * mm, 26 * mm, 8 * mm, 16 * mm, 26 * mm, 16 * mm, 25 * mm], st))
    story.append(PageBreak())

# ---- 조립 ----
by_field = {f: [] for f in FIELD_ORDER}
for k in sorted(demands, key=int):
    by_field.setdefault(field6t.get(k, "융합"), []).append(k)

story.append(NextPageTemplate("cover"))
cover()
intro_toc(by_field)
no = 0
for f in FIELD_ORDER:
    if not by_field[f]: continue
    no += 1
    chapter(no, f, by_field[f])

Doc(OUT).build(story)
print("saved:", OUT)
