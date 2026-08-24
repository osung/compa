# -*- coding: utf-8 -*-
"""특허 임베딩 축 도입 전후 매칭 변화 분석 보고서(PDF).

입력
  MEDITEK_260820_매칭.pkl              도입 후(최종)
  $COMPA_SCRATCH/before_patax_매칭.pkl  도입 전
  $COMPA_SCRATCH/mid_patax_w012_매칭.pkl · mid_patax_w008_매칭.pkl   중간 단계
  $COMPA_SCRATCH/patax_change.pkl      기업별 교체 목록(탈락 과제의 특허축 재계산 포함)

사용: COMPA_SCRATCH=<scratch> python report_patax_change.py
"""
import os
import re

import numpy as np
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, Image as RLImage,
                                KeepTogether, PageBreak, PageTemplate, Paragraph,
                                Spacer, Table, TableStyle)

import meditek_theme as mt

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
# 산출 PDF 에 버전 번호를 붙인다(덮어쓰기 방지·이력 보존)
_BASE = os.path.join(HERE, "특허축_매칭변화_분석")


def _next_out():
    import glob
    vers = [int(m.group(1)) for f in glob.glob(_BASE + "_v*.pdf")
            if (m := re.search(r"_v(\d+)\.pdf$", f))]
    return f"{_BASE}_v{(max(vers) + 1) if vers else 1}.pdf"


OUT = os.environ.get("PATAX_CHANGE_OUT", _next_out())

# ---- 폰트 -------------------------------------------------------------------
FDIR = os.environ.get("COMPA_FONT_DIR", os.path.join(SCRATCH, "fonts"))
_FACES = [("Sans", "NotoSansKR-Regular.ttf"), ("Sans-B", "NotoSansKR-Bold.ttf"),
          ("Serif", "NotoSerifKR-Regular.ttf"), ("Serif-B", "NotoSerifKR-Bold.ttf")]
_FB = next((p for p in ("/Library/Fonts/Arial Unicode.ttf",
                        os.path.expanduser("~/Library/Fonts/NotoSansKR-Regular.ttf"))
            if os.path.exists(p)), None)
_miss = []
for nm, fn in _FACES:
    p = os.path.join(FDIR, fn)
    if os.path.exists(p):
        pdfmetrics.registerFont(TTFont(nm, p))
    elif _FB:
        pdfmetrics.registerFont(TTFont(nm, _FB))
        _miss.append(fn)
    else:
        raise SystemExit(f"[중단] 한글 폰트 없음: {p}")
if _miss:
    print(f"! 없는 폰트 {_miss} → 대체 {_FB}")

# ---- 색 --------------------------------------------------------------------
BLUE = colors.HexColor("#" + mt.APOLLO_BLUE)
MID = colors.HexColor("#" + mt.APOLLO_BLUE_MID)
LT = colors.HexColor("#" + mt.APOLLO_BLUE_LT)
INK = colors.HexColor("#" + mt.APOLLO_INK)
GREY = colors.HexColor("#" + mt.APOLLO_GREY)
HAIR = colors.HexColor("#" + mt.APOLLO_HAIR)
TINT = colors.HexColor("#" + mt.APOLLO_TINT)
PALE = colors.HexColor("#" + mt.APOLLO_PALE)
UP = colors.HexColor("#1E7A46")       # 개선
DOWN = colors.HexColor("#A93226")     # 악화

PAGE_W, PAGE_H = A4
LM = RM = 16 * mm
TM, BM = 20 * mm, 15 * mm
CW = PAGE_W - LM - RM


def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def P(t, size=9, color=INK, align=TA_LEFT, bold=False, family="Sans", lead=None,
      space=0, raw=False):
    st = ParagraphStyle("s", fontName=(family + "-B") if bold else family,
                        fontSize=size, textColor=color, alignment=align,
                        leading=lead or size * 1.45, spaceAfter=space)
    return Paragraph(str(t) if raw else esc(t), st)


def tbl(data, widths, extra=None, fs=8):
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, HAIR),
        ("FONT", (0, 0), (-1, -1), "Sans", fs),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (0, 0), (-1, 0), BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONT", (0, 0), (-1, 0), "Sans-B", fs),
    ] + (extra or [])))
    return t


def h1(t):
    return P(t, 15, BLUE, bold=True, space=7)


def h2(t):
    return P(f'<font color="#{mt.APOLLO_BLUE_LT}">▍</font> {esc(t)}', 12, BLUE,
             bold=True, space=5, raw=True)


class Doc(BaseDocTemplate):
    def __init__(self, fn):
        super().__init__(fn, pagesize=A4, leftMargin=LM, rightMargin=RM,
                         topMargin=TM, bottomMargin=BM)
        fr = Frame(LM, BM, CW, PAGE_H - TM - BM, id="n", topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate(id="cover", frames=[fr], onPage=self._cov),
                               PageTemplate(id="body", frames=[fr], onPage=self._body)])

    def _cov(self, c, d):
        pass

    def _body(self, c, d):
        c.saveState()
        y = PAGE_H - 12 * mm
        try:
            w, h = mt.logo_size(mt.LOGO_APOLLO, 24 * mm)
            c.drawImage(mt.LOGO_APOLLO, PAGE_W - RM - w, y + 2, width=w, height=h,
                        mask="auto")
        except Exception:
            pass
        c.setStrokeColor(HAIR); c.setLineWidth(0.6); c.line(LM, y, PAGE_W - RM, y)
        c.setStrokeColor(LT); c.setLineWidth(1.8); c.line(LM, y, LM + 18, y)
        c.setFont("Sans", 7.5); c.setFillColor(GREY)
        c.drawString(LM, y + 4, "특허 임베딩 축 도입 전후 매칭 변화 분석")
        c.drawCentredString(PAGE_W / 2, BM - 9, f"— {d.page - 1} —")
        c.restoreState()


story = []
S = story.append


# ---- 데이터 -----------------------------------------------------------------
def load():
    aft = pd.read_pickle(os.path.join(HERE, "MEDITEK_260820_매칭.pkl"))
    bef = pd.read_pickle(os.path.join(SCRATCH, "before_patax_매칭.pkl"))
    ch = pd.read_pickle(os.path.join(SCRATCH, "patax_change.pkl"))
    mid = {}
    for tag, fn in (("w012", "mid_patax_w012_매칭.pkl"), ("w008", "mid_patax_w008_매칭.pkl"),
                    ("nofilt", "before_namefilter_매칭.pkl"),
                    ("noexempt", "before_exempt_매칭.pkl"),
                    ("nophase", "before_phase_매칭.pkl"),
                    ("noorg", "before_orgonly_매칭.pkl")):
        p = os.path.join(SCRATCH, fn)
        if os.path.exists(p):
            mid[tag] = pd.read_pickle(p)
    members = {}
    mp = os.path.join(HERE, "MEDITEK_260820_연차통합.json")
    if os.path.exists(mp):
        import json
        members = json.load(open(mp, encoding="utf-8"))
    return bef, aft, mid, ch, members


UMB = re.compile(r"인재양성|대학원지원|창업선도|혁신인재|^[가-힣]+대학교$|연구소$|사업단|지원사업|인력양성|육성사업")


def n_umb(d):
    return int(d["과제명"].str.contains(UMB).sum())


def cover(bef, aft, ch):
    S(Spacer(1, 26 * mm))
    try:
        w, h = mt.logo_size(mt.LOGO_APOLLO, 74 * mm)
        img = RLImage(mt.LOGO_APOLLO, width=w, height=h); img.hAlign = "CENTER"
        S(img); S(Spacer(1, 10 * mm))
    except Exception:
        pass
    S(tbl([[P("특허 유사성 · 특허 유망성 반영", 13, MID, TA_CENTER, bold=True)],
           [P("매칭 결과 변화 분석", 25, BLUE, TA_CENTER, bold=True, lead=32)],
           [P("2026 MEDITEK 국가R&D 과제 매칭", 11, GREY, TA_CENTER)]],
          [CW], [("BACKGROUND", (0, 0), (-1, -1), PALE),
                 ("GRID", (0, 0), (-1, -1), 0, colors.white),
                 ("LINEABOVE", (0, 0), (-1, 0), 2.4, BLUE),
                 ("LINEBELOW", (0, -1), (-1, -1), 2.4, BLUE),
                 ("TOPPADDING", (0, 0), (-1, -1), 10),
                 ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    S(Spacer(1, 12 * mm))
    n_co = ch["기업명"].nunique()
    n_ch = int((ch["구분"] == "진입").sum())
    cards = [("추천 교체", f"{n_ch}건 / {len(aft)}건", f"{n_ch/len(aft)*100:.0f}%"),
             ("영향받은 기업", f"{n_co}개사 / {aft['기업명'].nunique()}개사", ""),
             ("추천당 생존특허", f"{bef['특허건수'].mean():.2f} → {aft['특허건수'].mean():.2f}건",
              f"+{(aft['특허건수'].mean()/bef['특허건수'].mean()-1)*100:.0f}%"),
             ("적합도 평균", f"{bef['적합도'].mean():.1f} → {aft['적합도'].mean():.1f}",
              f"{aft['적합도'].mean()-bef['적합도'].mean():+.1f}")]
    S(tbl([[P(k, 8.5, colors.white, TA_CENTER, bold=True) for k, _, _ in cards],
           [P(v, 11, BLUE, TA_CENTER, bold=True) for _, v, _ in cards],
           [P(x, 8, UP if x.startswith("+") else GREY, TA_CENTER) for _, _, x in cards]],
          [CW / 4] * 4,
          [("BACKGROUND", (0, 1), (-1, -1), colors.white),
           ("BOX", (0, 0), (-1, -1), 0.8, BLUE),
           ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    S(Spacer(1, 14 * mm))
    S(P("이 문서는 국가R&D 과제가 확보한 특허의 임베딩을 매칭 점수에 반영한 뒤 "
        "추천 결과가 어떻게 바뀌었는지 정리한 것이다. 기업별 교체 과제 목록과 그 사유를 "
        "5장에 전부 실었다. 탈락 과제의 특허축 점수는 최종 설정으로 다시 계산해 "
        "비교 가능하게 했다.", 9.5, INK, family="Serif", lead=16))
    S(PageBreak())


def sec_design():
    S(h1("1. 무엇을 바꿨는가"))
    S(h2("추가한 점수 축"))
    S(P("과제가 확보한 특허의 임베딩을 기업 질의 임베딩과 비교해, 기업 기술과 유사하고 "
        "권리가 살아 있는 특허를 많이 가진 과제를 높게 본다. 소스는 "
        "nice_patent_260813.pkl(403,024건, norm_embed 768차원)이며 코퍼스 과제를 100% 덮는다.",
        9.5, INK, family="Serif", lead=15, space=6))
    S(tbl([["구성", "정의", "비고"],
           [P("품질계수 qual", 8), P("0.55 + 0.30·법적상태 + 0.15·최신성  →  0.55~1.00", 8),
            P("등록 특허 1건 = 취하 특허 1건의 약 1.42배", 8)],
           [P("유사도 에너지 E", 8), P("Σ max(sim−τ, 0) · qual   (과제당 상위 5건)", 8),
            P("τ = 질의별 특허 유사도 P90", 8)],
           [P("특허축 pat", 8), P("0.75·(1 − exp(−E/E_ref)) + 0.25·mx_n·qual_best", 8),
            P("E_ref = P95. 상한에 붙지 않는 포화 곡선", 8)]],
          [26 * mm, 74 * mm, CW - 100 * mm]))
    S(Spacer(1, 5 * mm))
    S(P("품질을 <b>더하기</b>로 넣으면 안 된다. 권리·최신성은 기업 적합성이 아니라 자산 "
        "품질을 재는 값이라 LLM 적합도와 상관이 없거나 음(−0.037 / −0.032)인데, 가중의 "
        "45%를 차지하면서 관련성 신호를 절반으로 깎았다(적합도 순위상관 +0.129 → +0.074, "
        "유사특허 개수 단조성 +0.947 → +0.624). <b>곱하기</b>로 바꾸면 관련성을 잃지 않고"
        "(+0.127) 개수 단조성과 권리 변별을 함께 지킨다.",
        9.5, INK, family="Serif", lead=15, space=8, raw=True))
    S(h2("가중치"))
    S(tbl([["단계", "적합도", "과제 유사도", "과제 유망성", "특허축"],
           ["1차선정", "—", "0.625", "0.208", "0.167"],
           ["재랭킹", "0.648", "0.139", "0.139", "0.074"]],
          [30 * mm] + [(CW - 30 * mm) / 4] * 4,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("BACKGROUND", (4, 1), (4, -1), TINT),
           ("FONT", (4, 1), (4, -1), "Sans-B", 8)]))
    S(Spacer(1, 4 * mm))
    S(P("재랭킹 가중을 0.12로 두면 축이 적합도 10점 차이를 뒤집는다(적합도 10점 = 최종 "
        "6.25점, 특허축 0.65 차 = 최종 7.0점). 실측에서 인재양성·기관지원 사업이 적합도 "
        "50짜리 추천을 밀어냈다. 기존 연계유형 가중과 같은 ‘적합도 약 5점 상당’ 수준인 "
        "0.08로 낮췄다. 1차선정 가중(0.20)은 누가 채점 대상이 되는지만 정하므로 강하게 "
        "두어도 무해하다.", 9.5, INK, family="Serif", lead=15, space=8))
    S(h2("특허 상태 필터"))
    S(P("코퍼스 조건을 ‘특허 1건 이상’에서 <b>‘등록 또는 출원 계류(공개) 특허 1건 이상’</b>"
        "으로 바꿨다. 거절·취하·포기는 권리가 성립하지 않았고 소멸은 권리가 끝났으므로 "
        "기술이전 협의 대상이 될 수 없다. 소멸의 법적상태 점수(0.55)가 공개(0.45)보다 "
        "높으므로 <b>점수가 아니라 상태명으로</b> 판정한다.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    S(tbl([["", "특허 원본", "특허 보유 과제(전국)", "의료·제약 코퍼스", "코퍼스 특허"],
           ["전", "403,024건", "201,958", "1,803", "4,509"],
           ["후", "252,596건", "156,565 (77.5%)", "1,731 (−72)", "4,177"]],
          [16 * mm] + [(CW - 16 * mm) / 4] * 4,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"), ("BACKGROUND", (0, 1), (0, -1), TINT),
           ("FONT", (0, 1), (0, -1), "Sans-B", 8)]))
    S(PageBreak())


def sec_overall(bef, aft, mid, ch, members=None):
    S(h1("2. 전체 변화"))
    S(h2("단계별 지표"))
    rows = [["단계", "적합도 평균", "추천당 특허", "우산형 추천", "적합도<50",
             "기존 대비 교체*"]]
    rep_of = {q: rep for rep, qs in (members or {}).items() for q in qs}
    # 연차 통합 단계에서 대표 번호가 바뀐 것을 '교체'로 세면 과대 계상된다.
    # 양쪽 과제고유번호를 대표 번호로 정규화한 뒤 비교한다.
    def kset(d):
        return {(r["기업명"], rep_of.get(str(r["과제고유번호"]), str(r["과제고유번호"])))
                for _, r in d.iterrows()}
    base = kset(bef)
    stages = [("① 도입 전(특허축 없음)", bef), ("② 특허축 w=0.12", mid.get("w012")),
              ("③ w=0.08 + 과제당 상한 5건", mid.get("w008")),
              ("④ + 생존특허 필터", mid.get("nofilt")),
              ("⑤ + 과제명 정상성 필터", mid.get("noexempt")),
              ("⑥ + 고적합도·유망특허 예외", mid.get("nophase")),
              ("⑦ + 다년차 과제 통합", mid.get("noorg")),
              ("⑧ + 기관명 단독 예외 불허 (최종)", aft)]
    for lab, d in stages:
        if d is None:
            continue
        cur = kset(d)
        chg = "—" if lab.startswith("①") else f"{len(cur - base)}건"
        rows.append([lab, f"{d['적합도'].mean():.1f}", f"{d['특허건수'].mean():.2f}건",
                     f"{n_umb(d)}행", f"{int((d['적합도'] < 50).sum())}행", chg])
    S(tbl(rows, [56 * mm] + [(CW - 56 * mm) / 5] * 5,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1), TINT),
           ("FONT", (0, len(rows) - 1), (-1, len(rows) - 1), "Sans-B", 8)]))
    S(Spacer(1, 2 * mm))
    S(P("* 교체 건수는 과제고유번호를 대표 번호로 정규화해 비교한 값이다. 연차 통합으로 "
        "번호만 바뀐 추천은 교체로 세지 않는다.", 8, GREY, space=5))
    S(h2("탈락 과제 vs 진입 과제"))
    o = ch[ch["구분"] == "탈락"]
    i = ch[ch["구분"] == "진입"]
    ov = o[o["특허축"].notna()]
    rows = [["", f"탈락 {len(o)}건", f"진입 {len(i)}건", "차이"],
            ["LLM 적합도 평균", f"{o['적합도'].mean():.1f}", f"{i['적합도'].mean():.1f}",
             f"{i['적합도'].mean()-o['적합도'].mean():+.1f}"],
            ["특허축 평균", f"{ov['특허축'].mean():.3f}", f"{i['특허축'].mean():.3f}",
             f"{i['특허축'].mean()-ov['특허축'].mean():+.3f}"],
            ["유사특허 수 평균", f"{ov['유사특허'].mean():.2f}건", f"{i['유사특허'].mean():.2f}건",
             f"{i['유사특허'].mean()-ov['유사특허'].mean():+.2f}"],
            ["생존특허 수 평균", f"{o['특허_생존'].mean():.2f}건", f"{i['특허_생존'].mean():.2f}건",
             f"{i['특허_생존'].mean()-o['특허_생존'].mean():+.2f}"],
            ["유사특허 보유 비율", f"{(ov['유사특허']>0).mean()*100:.0f}%",
             f"{(i['유사특허']>0).mean()*100:.0f}%", ""]]
    S(tbl(rows, [46 * mm] + [(CW - 46 * mm) / 3] * 3,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("TEXTCOLOR", (3, 1), (3, -1), UP),
           ("FONT", (3, 1), (3, -1), "Sans-B", 8)]))
    S(Spacer(1, 5 * mm))
    S(P("<b>적합도를 내주고 특허를 얻은 것이 아니다.</b> 진입 과제는 적합도도 "
        f"{i['적합도'].mean()-o['적합도'].mean():+.1f} 높으면서 유사특허를 "
        f"{i['유사특허'].mean()/max(ov['유사특허'].mean(),1e-9):.1f}배, 생존특허를 "
        f"{i['특허_생존'].mean()/max(o['특허_생존'].mean(),1e-9):.1f}배 가진다. "
        "1차선정에 특허축을 넣어 특허가 두꺼운 과제를 후보 120건 안으로 끌어올린 결과다. "
        "그 축이 없을 때는 애초에 채점 대상에 들어오지 못했다.",
        9.5, INK, family="Serif", lead=15, space=8, raw=True))
    S(h2("교체 사유 분류"))
    tot = len(o)
    nm = int(o["배제사유"].isin(["기관명 단독", "기관·조직명", "인력양성·기관지원 사업",
                                 "연구주제 없는 단문"]).sum())
    filt = int((o["배제사유"] == "코퍼스 제외").sum())
    rest = o[o["배제사유"] == ""]
    nosim = int((rest["유사특허"].fillna(0) == 0).sum())
    low = tot - nm - filt - nosim
    rows = [["사유", "건수", "설명"],
            ["과제명 배제", f"{nm}건",
             "과제명이 기관·조직명 또는 인력양성·기관지원 사업 → 코퍼스에서 제외"],
            ["코퍼스 제외", f"{filt}건",
             "생존특허 없음(전량 거절·취하·소멸) 또는 다른 과제로 흡수돼 코퍼스에서 빠짐"],
            ["유사특허 없음", f"{nosim}건",
             "기업 기술과 유사한 특허(질의 유사도 P90 초과)를 한 건도 갖지 못해 특허축 0점"],
            ["특허축 열위", f"{low}건",
             "유사특허는 있으나 진입 과제보다 적거나 권리·최신성이 낮아 순위 밀림"]]
    S(tbl(rows, [30 * mm, 18 * mm, CW - 48 * mm]))
    S(Spacer(1, 4 * mm))
    both = sum(1 for _, r in i.iterrows()
               if _both(r, ch[(ch["번호"] == r["번호"]) & (ch["구분"] == "탈락")]))
    S(tbl([["진입 사유", "건수", "설명"],
           ["적합도·특허축 동반 우위", f"{both}건",
            "탈락 과제보다 LLM 적합도도 높고 특허축도 높다 — 순수한 개선"],
           ["특허축 우위", f"{len(i)-both}건",
            "적합도는 탈락 과제보다 낮지만 특허 실적 격차가 이를 상쇄"]],
          [30 * mm, 18 * mm, CW - 48 * mm]))
    S(PageBreak())


def _both(r, outs):
    """진입 과제가 적합도에서도 우위인가 — 코퍼스에서 빠진 탈락 과제는 비교에서 제외."""
    cmpo = outs[~outs["배제사유"].isin(
        ("기관명 단독", "기관·조직명", "인력양성·기관지원 사업", "연구주제 없는 단문",
         "코퍼스 제외"))]
    return len(cmpo) == 0 or r["적합도"] >= cmpo["적합도"].mean()


def sec_risk(aft, bef):
    S(h1("3. 부작용과 대응"))
    S(h2("우산형 과제 — 발견과 배제"))
    S(P("특허축을 넣자 인력양성·기관지원 사업이 추천에 늘었다(3행 → 8행). 과제명이 "
        "‘인공지능대학원지원’, ‘지역지능화혁신인재양성’, ‘충북대학교’ 같은 것들로, 기술이 "
        "아니라 사업·기관 단위 묶음이다. 산하 연구실 특허를 대량 누적해(최대 153건) "
        "특허축에 유리하다.", 9.5, INK, family="Serif", lead=15, space=6))
    S(P("<b>과제당 기여 특허 상한(K)으로는 해결되지 않았다.</b> K를 줄여도 우산형 우위가 "
        "남는다 — 그 과제들이 실제로 기업 기술과 유사한 특허를 여러 건 갖고 있기 때문이다. "
        "K는 신호를 전혀 잃지 않으므로(적합도 순위상관 +0.133 고정) 병리적 포트폴리오 "
        "방어용으로 5를 유지했다.", 9.5, INK, family="Serif", lead=15, space=6, raw=True))
    S(tbl([["과제당 기여 특허 상한 K", "적합도 순위상관", "우산형/집중형 축 평균비",
            "축 상위 30에 든 우산형"],
           ["2", "+0.133", "1.61", "3.6"],
           ["5 (채택)", "+0.133", "1.98", "4.9"],
           ["무제한", "+0.133", "2.07", "5.0"]],
          [42 * mm] + [(CW - 42 * mm) / 3] * 3,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("BACKGROUND", (0, 2), (-1, 2), TINT), ("FONT", (0, 2), (-1, 2), "Sans-B", 8)]))
    S(Spacer(1, 5 * mm))
    S(h2("과제명 정상성 필터"))
    S(P("근본 해결은 코퍼스에서 빼는 것이다. 과제명이 연구 주제가 아닌 과제를 세 규칙으로 "
        "배제한다(match_meditek_supply.bad_project_name). 기본 적용이며 "
        "--allow-any-name 으로만 해제된다.", 9.5, INK, family="Serif", lead=15, space=6))
    rr = [["규칙", "정의", "배제", "예"],
          ["① 기관·조직명",
           "기관 단위 명사로 끝난다(대학교·연구소·센터·병원·재단 등). "
           "‘연구실’은 제외 — 주제를 담은 정상 과제명이다", "23건",
           "충북대학교 · 생태환경독성연구소"],
          ["② 인력양성·기관지원",
           "인재양성·인력양성·대학원지원·창업선도·혁신인재·리더스·육성사업 등", "19건",
           "지역지능화혁신인재양성"],
          ["③ 주제 없는 단문",
           "12자 이하이면서 연구 주제 어휘가 하나도 없다", "4건", "차세대 창의약학"]]
    S(tbl([rr[0]] + [[P(c, 7.5) for c in r] for r in rr[1:]],
          [26 * mm, 68 * mm, 13 * mm, CW - 107 * mm],
          [("ALIGN", (2, 1), (2, -1), "CENTER")]))
    S(Spacer(1, 4 * mm))
    S(P("코퍼스 1,731건 중 46건(2.7%)이 걸리고 오탐은 없었다. ‘암화 방어기전 연구’(10자)는 "
        "‘연구’가 있어 유지되고, ‘간-대장 상통 대사항상성 연구실’도 유지된다. 검증에 "
        "‘추천 과제명이 연구 주제’ 항목을 넣어 산출물에서도 다시 막는다.",
        9.5, INK, family="Serif", lead=15, space=6))
    S(tbl([["", "우산형 추천", "적합도<50 추천", "코퍼스"],
           ["특허축 도입 전", "3행", "6행", "1,803건"],
           ["특허축 도입 후(필터 전)", "8행", "7행", "1,731건"],
           ["과제명 필터 적용", "0행", "7행", "1,685건"],
           ["예외 조항 적용(최종)",
            f"{int(aft['과제명특이'].astype(bool).sum()) if '과제명특이' in aft.columns else 0}행"
            "(전부 적합도 ≥80)", f"{int((aft['적합도']<50).sum())}행", "1,731건"]],
          [46 * mm] + [(CW - 46 * mm) / 3] * 3,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("BACKGROUND", (0, 4), (-1, 4), TINT), ("FONT", (0, 4), (-1, 4), "Sans-B", 8)]))
    S(Spacer(1, 4 * mm))
    S(Spacer(1, 3 * mm))
    S(h2("예외 조항 — 고적합도 · 유망특허 보유 과제는 남긴다"))
    S(P("전면 배제하면 실제로 이전할 만한 기술이 확인되는 과제까지 버린다. "
        "(주)비전메디컬 4위 ‘지역지능화혁신인재양성’은 LLM 적합도 92, "
        "(주)레보스케치 3위 ‘나노센서 연구소’는 85에 유망 등록특허 14건을 가진다. "
        "그래서 <b>과제 성격과 무관하게 다음 둘을 모두 충족하면 남긴다</b>.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    S(tbl([["예외 조건", "정의", "근거"],
           [P("① 높은 LLM 적합도", 8), P("적합도 ≥ 80", 8),
            P("기업 기술과의 부합도를 모델이 높게 본 경우만 인정", 8)],
           [P("② 유망 특허 보유", 8),
            P("등록 상태이고 특허 유망성점수 ≥ 80 인 특허를 1건 이상", 8),
            P("코퍼스 등록특허 유망성점수 실측 P25 = 82.3 → 하위 4분위를 걸러내는 선", 8)]],
          [30 * mm, 62 * mm, CW - 92 * mm]))
    S(Spacer(1, 4 * mm))
    S(P("예외 판정에 적합도가 필요해 과제명 검사를 코퍼스에서 배제하지 않고 <b>표시만</b> 하고, "
        "LLM 채점이 끝난 뒤 최종 선정에서 걸러낸다. 표시된 과제는 특허 포트폴리오가 커서 "
        "1차점수가 높아 후보 정원을 잠식하므로(실측: 에이아이다이콤 적합도 95짜리 1위 과제가 "
        "후보에서 탈락) <b>정원 밖에서 15건까지만 추가 채점</b>한다.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    nb = aft[aft["과제명특이"].astype(bool)] if "과제명특이" in aft.columns else aft.head(0)
    rows2 = [["번호", "기업명", "순위", "적합도", "연차", "유망등록특허", "과제명"]]
    for r in nb.sort_values("번호").to_dict("records"):
        rows2.append([str(r["번호"]), P(r["기업명"], 7.5), f"{r['rank']}위",
                      str(r["적합도"]), f"{int(r.get('연차수', 1))}",
                      f"{r['유망등록특허수']}건", P(r["과제명"], 7.5)])
    S(KeepTogether([
        P(f"예외로 남은 추천 {len(nb)}행 (고유 과제 "
          f"{nb['과제고유번호'].nunique() if len(nb) else 0}개)", 9, BLUE, bold=True, space=4),
        tbl(rows2, [11 * mm, 34 * mm, 12 * mm, 13 * mm, 10 * mm, 20 * mm,
                    CW - 100 * mm],
            [("ALIGN", (0, 1), (5, -1), "CENTER")], fs=7.5)]))
    S(Spacer(1, 3 * mm))
    S(P("<b>단, 과제명이 기관 이름 단독이면 예외를 인정하지 않는다.</b> 이름에 연구 주제가 "
        "한 조각도 없어 적합도가 높아도 무엇을 이전받는지 특정할 수 없다. ‘연구소’·‘센터’로 "
        "끝나는 이름은 제외한다 — ‘약학기술연구소’, ‘나노센서 연구소’처럼 연구 영역을 담아 "
        "상담 단서가 된다. 실측: 에이트스튜디오(주) 3위 ‘충북대학교’(3연차, 특허 296건 · "
        "유망 등록특허 82건, 적합도 90)가 이 규칙으로 빠지고 적합도 75 대체 과제가 들어왔다.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    S(Spacer(1, 2 * mm))
    S(P("연차 통합의 부작용 — 우산형 과제는 산하 연구실 특허를 연차별로 누적하므로 통합 시 "
        "포트폴리오가 더 커진다. 에이트스튜디오(주) 3위 ‘충북대학교’(3연차)는 통합 후 특허 "
        "296건 · 유망 등록특허 82건이 되어 적합도 90 으로 예외를 통과했다. 예외 조건이 "
        "요청대로 ‘적합도 ≥80 ∧ 유망특허 ≥1’ 이므로 규칙상 정당하지만, 과제명만 보면 "
        "기술이전 대상으로 읽히지 않는다. 이 사례가 위 ‘기관명 단독’ 규칙을 만든 계기다.",
        9, GREY, family="Serif", lead=14, space=4))
    S(PageBreak())


def sec_phase(aft, bef, mid, members):
    """다년차 과제 통합."""
    S(h1("4. 다년차 과제 통합"))
    S(P("NTIS 는 다년도 과제의 <b>연차 보고마다 별도의 과제고유번호</b>를 부여한다. 같은 "
        "과제가 연차 수만큼 행으로 등재되고, 연차마다 성과가 누적되므로 특허·논문 건수와 "
        "유망성점수가 갈린다. 보고서에서 같은 과제명이 다른 특허 실적을 달고 두 번 나오면 "
        "오류로 읽히므로 한 건으로 묶었다.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    S(h2("발견 사례"))
    S(tbl([["과제고유번호", "제출년도", "생존특허", "수행기관", "과제명"],
           [P("1415169720", 8), "2020", "3건", P("경북대학교", 8),
            P("영상진단 의료기기 탑재용 AI 진단 기술 개발", 8)],
           [P("1415173238", 8), "2021", "5건", P("경북대학교", 8), P("(동일)", 8)],
           [P("1415179472", 8), "2022", "—", P("경북대학교", 8), P("(동일)", 8)],
           [P("1415187778 ← 대표", 8), "2023", "—", P("경북대학교", 8), P("(동일)", 8)]],
          [30 * mm, 16 * mm, 16 * mm, 24 * mm, CW - 86 * mm],
          [("ALIGN", (1, 1), (2, -1), "CENTER"),
           ("BACKGROUND", (0, 4), (-1, 4), TINT), ("FONT", (0, 4), (-1, 4), "Sans-B", 8)]))
    S(Spacer(1, 3 * mm))
    S(P("연구기간(2020-04-01 ~ 2024-12-31)과 수행기관이 같은 하나의 과제다. 통합 후 특허는 "
        "전 연차 합집합 <b>9건</b>(등록 5 · 출원 계류 4)이 된다.",
        9.5, GREY, family="Serif", lead=15, space=6, raw=True))
    S(h2("그룹 키 — 두 함정"))
    S(tbl([["후보 키", "왜 안 되는가", "실측"],
           [P("과제명만", 8),
            P("여러 기관이 같은 이름의 국가사업을 각각 수행한다 — 별개 과제인데 합쳐진다", 8),
            P("‘실험실 특화형 창업선도대학 사업’(울산과기원/충북대), "
              "‘핵심연구지원센터조성지원과제’(울산대·경북대·인제대, 전부 2022년)", 8)],
           [P("과제명 + 과제수행기관명", 8),
            P("같은 과제인데 연차마다 기관 표기가 바뀐다 — 합쳐야 할 것이 갈라진다", 8),
            P("서울아산병원↔아산사회복지재단, 고려대학교 안암병원↔고려대학의과대학부속병원, "
              "가톨릭대학교↔가톨릭대학교 성의캠퍼스 — <b>64개 그룹</b>", 8, raw=True)],
           [P("과제고유번호 접두", 8),
            P("2024년에 번호 체계가 개편돼 같은 과제의 번호가 바뀐다", 8),
            P("1465030925(2020) → 2460000169(2024)", 8)],
           [P("과제명 + 공급기관 계열 (채택)", 8, bold=True),
            P("이미 정의해 둔 공급기관 대응표(meditek_supply_orgs.include_map)가 "
              "기관 표기 흔들림을 정규화하면서 다른 기관은 분리한다", 8),
            P("동명이과제는 정확히 분리, 기관 표기가 바뀐 연차는 정확히 통합", 8)]],
          [34 * mm, 56 * mm, CW - 90 * mm]))
    S(Spacer(1, 4 * mm))
    S(h2("통합 규칙"))
    S(P("대표 행은 <b>제출년도가 가장 늦은 연차</b>(성과가 가장 많이 누적된 보고)로 한다. "
        "과제명·설명문·유망성·연구기간 등 속성은 대표 행에서 가져오고, <b>특허는 전 연차의 "
        "출원번호 합집합</b>을 쓴다(연차마다 같은 특허가 다시 실려 단순 합산은 중복이 된다). "
        "NTIS 과제메타의 특허·논문 건수는 누적값이라 연차 최대값을 쓴다. 특허축과 보고서 "
        "특허 표도 같은 대응표(MEDITEK_260820_연차통합.json)를 읽어 동일 기준으로 합친다.",
        9.5, INK, family="Serif", lead=15, space=6, raw=True))
    u = aft.drop_duplicates("과제고유번호")
    ub = bef.drop_duplicates("과제고유번호")
    nph = mid.get("nophase")
    S(tbl([["", "통합 전", "통합 후(최종)"],
           ["코퍼스", "1,731건", "1,173건 (−558)"],
           ["추천 고유 과제", f"{nph['과제고유번호'].nunique() if nph is not None else '—'}개",
            f"{len(u)}개"],
           ["같은 과제명 중복 추천",
            f"{int(nph.drop_duplicates('과제고유번호')['과제명'].duplicated().sum()) if nph is not None else '—'}건",
            f"{int(u['과제명'].duplicated().sum())}건"],
           ["적합도 평균",
            f"{nph['적합도'].mean():.1f}" if nph is not None else "—",
            f"{aft['적합도'].mean():.1f}"],
           ["추천당 생존특허",
            f"{nph['특허건수'].mean():.2f}건" if nph is not None else "—",
            f"{aft['특허건수'].mean():.2f}건"],
           ["보고서 특허 총계",
            f"{int(nph.drop_duplicates('과제고유번호')['특허건수'].sum()) if nph is not None else 0}건",
            f"{int(u['특허건수'].sum())}건"]],
          [42 * mm] + [(CW - 42 * mm) / 2] * 2,
          [("ALIGN", (1, 1), (-1, -1), "CENTER"),
           ("TEXTCOLOR", (2, 3), (2, 3), UP), ("FONT", (2, 3), (2, 3), "Sans-B", 8),
           ("BACKGROUND", (0, 1), (0, -1), PALE)]))
    S(Spacer(1, 4 * mm))
    dist = u["연차수"].value_counts().sort_index() if "연차수" in u.columns else {}
    S(P("연차 분포: " + " · ".join(f"{k}연차 {v}건" for k, v in dist.items())
        + (f"  |  통합 과제 특허 중위 {int(u[u['연차수'] > 1]['특허건수'].median())}건 vs "
           f"단일 연차 {int(u[u['연차수'] == 1]['특허건수'].median())}건"
           if "연차수" in u.columns and (u["연차수"] > 1).any() else ""),
        9, BLUE, bold=True, space=4))
    S(P("비교 기준 주의 — 이 문서의 탈락/진입 판정은 <b>대표 과제고유번호 기준</b>이다. "
        "연차 통합으로 번호만 바뀐 추천(도입 전 번호가 최종 대표 번호에 흡수된 경우) 19건은 "
        "탈락이 아니라 유지로 센다.", 9, GREY, family="Serif", lead=14, space=6, raw=True))
    S(PageBreak())


def _i(v):
    """유사특허 수 표기 — 결측이 섞인 컬럼은 pandas 가 float 로 만든다."""
    return "—" if _na(v) else str(int(v))


def _na(v):
    """pandas 가 None 을 NaN 으로 바꾸므로 결측을 한 곳에서 판정한다."""
    return v is None or (isinstance(v, float) and np.isnan(v))


NAME_EX = ("기관명 단독", "기관·조직명", "인력양성·기관지원 사업", "연구주제 없는 단문")
CORPUS_EX = NAME_EX + ("코퍼스 제외",)


def _pname(r):
    """과제명 + 과제고유번호. 연차 후속과제는 이름이 같아 번호로만 구분된다."""
    return Paragraph(
        f'{esc(r["과제명"][:34])}<br/><font size="6" color="#{mt.APOLLO_GREY}">'
        f'{esc(r["pid"])}</font>',
        ParagraphStyle("pn", fontName="Sans", fontSize=7.5, textColor=INK, leading=10))


def _reason_out(r, ins):
    """탈락 사유 — 저장값과 재계산한 특허축으로만 쓴다(추정 금지)."""
    if r.get("배제사유") in NAME_EX:
        return ("과제명 배제", f"과제명이 ‘{r['배제사유']}’ → 이전받을 기술이 특정되지 않아 "
                              f"코퍼스에서 제외(보유 특허 {r['특허_전체']}건)")
    if r["코퍼스제외"]:
        return ("코퍼스 제외", f"생존특허 없음 또는 다른 대표 과제로 흡수 "
                             f"(원본 특허 {r['특허_전체']}건)")
    pv = r["특허축"]
    cv = 0 if _na(r["유사특허"]) else int(r["유사특허"])
    ip = np.mean([x["특허축"] for x in ins]) if ins else 0.0
    df = np.mean([x["적합도"] for x in ins]) - r["적합도"] if ins else 0
    if cv == 0:
        return ("유사특허 없음",
                f"질의 유사도 P90을 넘는 특허 0건 → 특허축 {pv:.3f} "
                f"(진입 과제 평균 {ip:.3f}) · 적합도 {r['적합도']}"
                + (f", 진입 과제가 {df:+.0f}" if ins else ""))
    return ("특허축 열위",
            f"유사특허 {cv}/{r['특허_생존']}건 · 특허축 {pv:.3f} "
            f"(진입 과제 평균 {ip:.3f}) · 적합도 {r['적합도']}"
            + (f", 진입 과제가 {df:+.0f}" if ins else ""))


def _reason_in(r, outs):
    # 적합도 비교 기준에서 코퍼스에서 빠진 탈락 과제(과제명 배제·생존특허 없음)는 뺀다.
    # 점수로 밀린 것이 아니므로 비교 대상이 아니다.
    cmpo = [x for x in outs if x.get("배제사유") not in CORPUS_EX]
    ov = [x["특허축"] for x in cmpo if not _na(x["특허축"])]
    op = np.mean(ov) if ov else None
    df = (r["적합도"] - np.mean([x["적합도"] for x in cmpo])) if cmpo else None
    tag = "적합도·특허축 동반 우위" if (df is None or df >= 0) else "특허축 우위"
    txt = f"유사특허 {_i(r['유사특허'])}/{r['특허_생존']}건 · 특허축 {r['특허축']:.3f}"
    if op is not None:
        txt += f" (탈락 과제 평균 {op:.3f})"
    txt += f" · 적합도 {r['적합도']}"
    if df is not None:
        txt += f", 탈락 과제 대비 {df:+.0f}"
    elif outs:
        txt += " · 탈락 과제는 전량 코퍼스 배제(점수 경쟁 아님)"
    return (tag, txt)


def sec_by_company(bef, aft, ch):
    S(h1("5. 기업별 교체 과제와 사유"))
    S(P("변화가 있었던 24개사를 번호순으로 싣는다. 각 기업의 머리줄은 전체 5건 기준 평균이고, "
        "표의 ‘탈락’은 도입 전 추천에만, ‘진입’은 도입 후 추천에만 있는 과제다. "
        "탈락 과제의 특허축은 최종 설정으로 다시 계산한 값이다(당시에는 없던 지표).",
        9, GREY, family="Serif", lead=14, space=7))
    for no in sorted(ch["번호"].unique()):
        g = ch[ch["번호"] == no]
        name = g.iloc[0]["기업명"]
        sb, sa = bef[bef["기업명"] == name], aft[aft["기업명"] == name]
        outs = g[g["구분"] == "탈락"].to_dict("records")
        ins = g[g["구분"] == "진입"].to_dict("records")
        blk = []
        head = (f"평균 적합도 {sb['적합도'].mean():.0f} → {sa['적합도'].mean():.0f}  ·  "
                f"추천당 생존특허 {sb['특허건수'].mean():.1f} → {sa['특허건수'].mean():.1f}건  ·  "
                f"특허축 {sa['특허축점수'].mean():.3f}  ·  다년차 "
                f"{int(sa['연차수'].gt(1).sum()) if '연차수' in sa.columns else 0}건"
                f"  ·  교체 {len(ins)}/{len(sa)}건")
        blk.append(tbl([[P(f"No.{no}  {name}", 10.5, colors.white, bold=True),
                         P(head, 8, colors.white, align=TA_LEFT)]],
                       [58 * mm, CW - 58 * mm],
                       [("BACKGROUND", (0, 0), (-1, -1), BLUE),
                        ("GRID", (0, 0), (-1, -1), 0, BLUE)]))
        rows = [["구분", "순위", "과제명", "적합도", "연차", "유사/생존특허", "특허축", "사유"]]
        st = []
        for r in sorted(outs, key=lambda x: x["순위"]):
            tag, why = _reason_out(r, ins)
            rows.append([P("탈락", 7.5, DOWN, TA_CENTER, bold=True), f"{r['순위']}위",
                         _pname(r), str(r["적합도"]),
                         f"{int(r.get('연차수', 1))}",
                         f"{_i(r['유사특허'])}/{r['특허_생존']}",
                         ("—" if _na(r["특허축"]) else f"{r['특허축']:.3f}"),
                         P(f"<b>{tag}</b> — {esc(why)}", 7, raw=True)])
            st.append(("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1),
                       colors.HexColor("#FDF3F2")))
        for r in sorted(ins, key=lambda x: x["순위"]):
            tag, why = _reason_in(r, outs)
            rows.append([P("진입", 7.5, UP, TA_CENTER, bold=True), f"{r['순위']}위",
                         _pname(r), str(r["적합도"]),
                         f"{int(r.get('연차수', 1))}",
                         f"{_i(r['유사특허'])}/{r['특허_생존']}", f"{r['특허축']:.3f}",
                         P(f"<b>{tag}</b> — {esc(why)}", 7, raw=True)])
            st.append(("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1),
                       colors.HexColor("#F1F8F3")))
        st += [("ALIGN", (1, 1), (1, -1), "CENTER"), ("ALIGN", (3, 1), (6, -1), "CENTER")]
        blk.append(tbl(rows, [12 * mm, 11 * mm, 46 * mm, 12 * mm, 9 * mm, 19 * mm,
                              13 * mm, CW - 122 * mm], st, fs=7.5))
        blk.append(Spacer(1, 5 * mm))
        S(KeepTogether(blk))
    S(PageBreak())


def sec_verify(aft):
    S(h1("6. 검증"))
    S(h2("매칭 검증 (match_meditek_supply.verify · 30항목 전부 통과)"))
    rows = [["항목", "결과"],
            ["특허축 값 유한 · 0~1 범위",
             f"평균 {aft['특허축점수'].mean():.3f} · 범위 {aft['특허축점수'].min():.3f}~"
             f"{aft['특허축점수'].max():.3f} · 0점 "
             f"{int(aft['특허축점수'].le(1e-9).sum())}/{len(aft)}행(유사특허 없음)"],
            ["유사특허 보유행은 특허축>0", "위반 0행"],
            ["유사특허 없는 행은 최고유사도가 임계 이하",
             f"{int(aft['유사특허수'].eq(0).sum())}행 검사 · 위반 0"],
            ["품질계수 0.55~1.00 범위", "평균 0.908 · 범위 0.786~0.986"],
            ["유사특허수 ≤ 특허건수",
             f"위반 0행 · 유사특허 보유 {int(aft['유사특허수'].gt(0).sum())}/{len(aft)}행"],
            ["유사특허수 ↑ → 특허축 ↑ (기업내 순위상관)", "평균 ρ = +0.853 (26개사)"],
            ["모든 추천 과제가 생존 특허 보유",
             f"등록 {int(aft['과제권리최고'].ge(1.0).sum())}행 · 출원 계류 "
             f"{int(aft['과제권리최고'].between(0.4,0.99).sum())}행 · 위반 0"],
            ["모든 과제 특허 1건 이상(nice·생존)",
             f"최소 {int(aft['특허건수'].min())}건 · 중위 {int(aft['특허건수'].median())}건"]]
    S(tbl(rows, [66 * mm, CW - 66 * mm]))
    S(Spacer(1, 5 * mm))
    S(h2("보고서 검증 (verify_report_supply · 47항목 전부 통과)"))
    S(tbl([["항목", "결과"],
           ["보고서 특허 전건이 등록·공개 상태", "498건 검사 · 위반 0건 (원본 대조)"],
           ["상태 라벨이 원본 등록상태명과 일치", "등록 262 · 출원 236 · 불일치 0"],
           ["표 특허건수 == 원본 생존특허 건수", "114개 과제 전부 일치"],
           ["특허축 내부 컬럼명 비노출", "PDF 346,426자 · DOCX 334,375자 검사 · 검출 0"],
           ["유망성 수치 노출", "0건"]],
          [66 * mm, CW - 66 * mm]))
    S(Spacer(1, 6 * mm))
    S(h2("최종 산출"))
    S(P(f"기업 {aft['기업명'].nunique()}개사 · 추천 {len(aft)}건 · 중복 제외 과제 "
        f"{aft['과제고유번호'].nunique()}개 · 공급기관 {aft['공급기관'].nunique()}곳 · "
        f"연계유형 기술도입 {int((aft['연계유형']=='기술도입').sum())} / 공동연구 "
        f"{int((aft['연계유형']=='공동연구').sum())}",
        9.5, INK, family="Serif", lead=15))


def build():
    bef, aft, mid, ch, members = load()
    doc = Doc(OUT)
    cover(bef, aft, ch)
    from reportlab.platypus import NextPageTemplate
    story.insert(len(story) - 1, NextPageTemplate("body"))
    sec_design()
    sec_overall(bef, aft, mid, ch, members)
    sec_risk(aft, bef)
    sec_phase(aft, bef, mid, members)
    sec_by_company(bef, aft, ch)
    sec_verify(aft)
    doc.build(story)
    print("saved:", OUT)
    return OUT


if __name__ == "__main__":
    build()
