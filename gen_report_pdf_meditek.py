# -*- coding: utf-8 -*-
"""2026 MEDITEK 수요기술 매칭 보고서 PDF 생성(reportlab).

gen_report_pdf.py 의 MEDITEK 판. 조판·표·상세 블록은 gen_report_pdf 를 그대로 쓰고
표지/개요/목차/본문 골격만 교체한다(6T 장 구분 없음, COMPA 고정 문구 제거).
gen_report_meditek.py(docx) 와 문구·구성이 동일하다.

입력: MEDITEK_통합best.json, $COMPA_SCRATCH/{pid_fields,pid_patents}.json
폰트: $COMPA_FONT_DIR 또는 $COMPA_SCRATCH/fonts 의 NotoSansKR/NotoSerifKR 4종
     (없으면 gen_report_pdf 가 시스템 한글 폰트로 대체)

사용: COMPA_SCRATCH=<scratch> python gen_report_pdf_meditek.py
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# gen_report_pdf 는 import 시점에 데이터를 읽으므로 그 전에 입력을 지정해야 한다.
os.environ.setdefault("COMPA_REPORT_JSON", os.path.join(HERE, "MEDITEK_통합best.json"))

import gen_report_pdf as gp                                            # noqa: E402
from gen_report_pdf import (ACCENT, CW, HEADBG, HEADFG, INK, MUTED, NAVY,  # noqa: E402
                            HAIR, ZEBRA, DISCBD, DISCBG, P, esc, mktable,
                            base_grid, section_label, demand_block)
from reportlab.lib import colors                                       # noqa: E402
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT        # noqa: E402
from reportlab.lib.styles import ParagraphStyle                        # noqa: E402
from reportlab.lib.units import mm                                     # noqa: E402
from reportlab.platypus import (NextPageTemplate, PageBreak, Paragraph,  # noqa: E402
                                Spacer)

import gen_report_meditek as gm                                        # noqa: E402

gp.OMIT_INFO_FIELDS = gm.OMIT_INFO_FIELDS      # docx 와 동일한 정보표 구성 유지
gp.pidf = gp.pidf.get("fields", gp.pidf)       # {"name2pid":…,"fields":…} → pid 직접 조회 가능하게

OUT = os.environ.get("MEDITEK_PDF_OUT",
                     os.path.join(HERE, "MEDITEK_수요기술_매칭_보고서.pdf"))


def cover():
    s, d = gp.story, gp.demands
    n_dem = len(d)
    n_rec = sum(len(v["top5"]) for v in d.values())
    n_proj = len({t["과제고유번호"] for v in d.values() for t in v["top5"]})
    s.append(Spacer(1, 34 * mm))
    s.append(P(gm.COVER_EYEBROW, 10.5, ACCENT, TA_CENTER, bold=True, space=10))
    s.append(P(gm.COVER_TITLE1, 27, NAVY, TA_CENTER, bold=True, space=3))
    s.append(P(gm.COVER_TITLE2, 27, NAVY, TA_CENTER, bold=True, space=10))
    s.append(mktable([[""]], [70 * mm], [("LINEBELOW", (0, 0), (-1, -1), 1.5, NAVY)]))
    s.append(Spacer(1, 5 * mm))
    s.append(P(gm.COVER_SUB, 12, MUTED, TA_CENTER, space=22))

    meta = mktable([[P(f"{n_dem}건", 20, NAVY, TA_CENTER, bold=True),
                     P(f"{n_rec}건", 20, NAVY, TA_CENTER, bold=True),
                     P(f"{n_proj}개", 20, NAVY, TA_CENTER, bold=True)],
                    [P("대상 수요기술", 9.5, MUTED, TA_CENTER),
                     P("추천 과제", 9.5, MUTED, TA_CENTER),
                     P("중복 제외 과제", 9.5, MUTED, TA_CENTER)]],
                   [32 * mm, 32 * mm, 32 * mm],
                   [("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)])
    meta.hAlign = "CENTER"
    s.append(meta)
    s.append(Spacer(1, 12 * mm))
    s.append(P(f"발행일  {gm.PUBLISH_DATE}       {gm.ENGINE_NOTE}", 10, MUTED, TA_CENTER, space=16))
    s.append(mktable([[Paragraph(f'<font name="Sans-B" color="#A93226" size="10.5">※  유의사항</font><br/>'
                                 f'<font name="Sans" color="#7B241C" size="9.5">{esc(gp.DISCLAIMER)}</font>',
                                 ParagraphStyle("d", leading=14))]], [CW],
                     [("BOX", (0, 0), (-1, -1), 0.8, DISCBD), ("BACKGROUND", (0, 0), (-1, -1), DISCBG),
                      ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                      ("LEFTPADDING", (0, 0), (-1, -1), 11), ("RIGHTPADDING", (0, 0), (-1, -1), 11)]))
    s.append(NextPageTemplate("body"))
    s.append(PageBreak())


def intro_toc(ks):
    s, d = gp.story, gp.demands
    n_dem = len(ks)
    n_rec = sum(len(d[k]["top5"]) for k in ks)
    n_proj = len({t["과제고유번호"] for k in ks for t in d[k]["top5"]})
    s.append(section_label("개요", before=2))
    s.append(P(f"본 보고서는 2026 MEDITEK 참여기업이 제출한 수요기술 {n_dem}건을 대상으로, 공공 R&D 과제 "
               f"데이터베이스와의 의미 기반 매칭을 수행한 결과를 정리한 것이다. 각 수요기술에 대해 적합도가 높은 "
               f"추천 과제 상위 5건(총 {n_rec}건, 중복 제외 {n_proj}개 과제)을 선정하고, 매칭 근거와 상세 추천 "
               f"근거를 함께 제시하였다. {gm.CORPUS_NOTE}",
               10.5, INK, TA_JUSTIFY, leading=17, family="Serif", space=8))
    s.append(P("각 수요기술은 다음 순서로 구성된다.", 10.5, INK, family="Serif", space=3))
    for ln in ["수요 정보 — 기업명 · 수요기술 내용",
               "최종 추천 과제 Top 5 — 순위 · 과제명 · 수행기관 · 과제수행년도 · 매칭 근거",
               "추천 과제별 상세 정보표 및 상세 매칭 근거 — 연관성 · 수요기술 사양 적합성 · "
               "추천 과제의 우수성 · 유사 사례 및 실적"]:
        s.append(Paragraph(f'<font name="Sans-B" color="#0E7C86">· </font>'
                           f'<font name="Serif" color="#1B2430">{esc(ln)}</font>',
                           ParagraphStyle("b", fontSize=10, leading=15, leftIndent=14, spaceAfter=3)))

    s.append(section_label("목차", before=16))
    def pg(v): return str(v) if v else "··"
    rows = [[P("번호", 9.5, HEADFG, TA_CENTER, bold=True),
             P("기업명", 9.5, HEADFG, bold=True),
             P("면", 9.5, HEADFG, TA_RIGHT, bold=True)]]
    sty = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEBELOW", (0, 0), (-1, -1), 0.3, HAIR),
           ("BACKGROUND", (0, 0), (-1, 0), HEADBG),
           ("TOPPADDING", (0, 0), (-1, -1), 4.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5)]
    for k in ks:
        rows.append([P(f"수요 {k}", 9, ACCENT, TA_CENTER, bold=True),
                     P(d[k]["기업명"], 9, INK, leading=11.5),
                     P(pg(gp.PAGE_MAP.get(k)), 9, MUTED, TA_RIGHT)])
    s.append(mktable(rows, [17 * mm, CW - 45 * mm, 12 * mm], sty))
    s.append(PageBreak())


def demand_list(ks):
    """분야별 장 대신 본문 첫 페이지에 전체 수요기술 목록."""
    s, d = gp.story, gp.demands
    title = P("수요기술 목록", 21, NAVY, bold=True, space=1)
    title._chapter_no = 1                       # 목차 페이지 산출 마커(gen_report_pdf 규약)
    s.append(title)
    s.append(mktable([[P("TECHNOLOGY DEMANDS", 9, ACCENT, bold=True)]], [CW],
                     [("LINEBELOW", (0, 0), (-1, -1), 1.2, NAVY),
                      ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    s.append(Spacer(1, 3))
    s.append(P(f"수요기술 {len(ks)}건", 10, MUTED, space=10))
    s.append(section_label("수요기술 목록", before=2, after=6))
    rows = [[P(x, 9.5, HEADFG, TA_CENTER, bold=True) for x in ("번호", "수요기술명", "기업명")]]
    for k in ks:
        rows.append([P(k, 9.5, NAVY, TA_CENTER, bold=True),
                     P(d[k]["수요기술명"], 9.5, leading=12),
                     P(d[k]["기업명"], 9.5, INK, TA_CENTER)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)])
    for i in range(2, len(rows), 2):
        st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    s.append(mktable(rows, [14 * mm, CW - 54 * mm, 40 * mm], st))
    s.append(PageBreak())


def assemble():
    gp.story.clear()
    gp.story.append(NextPageTemplate("cover"))
    ks = sorted(gp.demands, key=int)
    cover()
    intro_toc(ks)
    demand_list(ks)
    for k in ks:
        demand_block(k, gp.demands[k], title=gp.demands[k]["기업명"])
    return list(gp.story)


if __name__ == "__main__":
    gp.assemble = assemble          # main() 이 모듈 전역에서 조회하므로 교체가 반영된다
    gp.main(OUT)
