# -*- coding: utf-8 -*-
"""2026 MEDITEK 기업 기술 × 공급기관 수행 국가 R&D 과제 매칭 보고서 PDF(reportlab).

gen_report_pdf.py 의 조판·과제 상세 블록을 그대로 쓰고 표지/개요/목차/요약/기업 블록만
교체한다. 문구·구성은 gen_report_meditek_supply.py(docx) 와 동일하다.

입력: MEDITEK_260820_보고서.json, $COMPA_SCRATCH/{pid_fields,pid_patents}.json
사용: COMPA_SCRATCH=<scratch> python gen_report_pdf_meditek_supply.py
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))

# gen_report_pdf 는 import 시점에 입력 JSON 을 읽으므로 그 전에 지정해야 한다.
os.environ.setdefault("COMPA_REPORT_JSON",
                      os.environ.get("MEDITEK_SUPPLY_JSON",
                                     os.path.join(HERE, "MEDITEK_260820_보고서.json")))

import gen_report_pdf as gp                                              # noqa: E402
from gen_report_pdf import (ACCENT, CW, HAIR, HEADBG, HEADFG, INK, LABELBG,  # noqa: E402
                            MUTED, NAVY, ZEBRA, DISCBD, DISCBG, P, base_grid, esc,
                            mktable, section_label)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT           # noqa: E402
from reportlab.lib.styles import ParagraphStyle                           # noqa: E402
from reportlab.lib.units import mm                                        # noqa: E402
from reportlab.platypus import (NextPageTemplate, PageBreak, Paragraph,   # noqa: E402
                                Spacer)

import gen_report_meditek_supply as gs                                    # noqa: E402
import gen_report_meditek_top10 as gm                                     # noqa: E402
import report_text_fix as tf                                              # noqa: E402

gp.OMIT_INFO_FIELDS = gm.gr.OMIT_INFO_FIELDS   # docx 와 동일한 정보표 구성(8개 항목 전부)
# 표기 교정 — docx 판과 같은 함수를 같은 대상에 적용해 두 산출물의 본문이 어긋나지 않게 한다
_FIXLOG = tf.fix_report(gp.demands, gp.patents, top_key=gs.TOP_KEY)
print(f"[표기 교정] {tf.report_stats()}")
gp.pidf = gp.pidf.get("fields", gp.pidf)       # {"name2pid":…,"fields":…} → pid 직접 조회
gp.DEMAND_TAG_FMT = "No.{no}"                  # 매칭 단위가 수요가 아니라 기업이다

OUT = os.environ.get("MEDITEK_SUPPLY_PDF_OUT",
                     os.path.join(HERE, "MEDITEK_공급기관_국가RnD_매칭보고서.pdf"))
TOP_KEY = gs.TOP_KEY


def npat(tp):
    """표에 싣는 특허 건수 — 상세 페이지의 특허 실적 목록과 같은 소스에서 센다."""
    return len(gp.patents.get(str(tp["과제고유번호"]), []))


def _counts(d):
    n_rec = sum(len(v[TOP_KEY]) for v in d.values())
    n_proj = len({t["과제고유번호"] for v in d.values() for t in v[TOP_KEY]})
    n_sup = len({t.get("공급기관", "") for v in d.values() for t in v[TOP_KEY]
                 if t.get("공급기관")})
    return n_rec, n_proj, n_sup


def cover():
    s, d = gp.story, gp.demands
    n_rec, n_proj, n_sup = _counts(d)
    s.append(Spacer(1, 34 * mm))
    s.append(P(gs.COVER_EYEBROW, 9.5, ACCENT, TA_CENTER, bold=True, space=10))
    s.append(P(gs.COVER_TITLE1, 27, NAVY, TA_CENTER, bold=True, space=3))
    s.append(P(gs.COVER_TITLE2, 27, NAVY, TA_CENTER, bold=True, space=10))
    s.append(mktable([[""]], [70 * mm], [("LINEBELOW", (0, 0), (-1, -1), 1.5, NAVY)]))
    s.append(Spacer(1, 5 * mm))
    s.append(P(gs.COVER_SUB, 12, MUTED, TA_CENTER, space=22))

    vals = [("대상 기업", f"{len(d)}개사"), ("추천 과제", f"{n_rec}건"),
            ("중복 제외 과제", f"{n_proj}개"), ("공급기관", f"{n_sup}곳")]
    meta = mktable([[P(v, 18, NAVY, TA_CENTER, bold=True) for _, v in vals],
                    [P(l, 9.5, MUTED, TA_CENTER) for l, _ in vals]],
                   [30 * mm] * 4,
                   [("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6)])
    meta.hAlign = "CENTER"
    s.append(meta)
    s.append(Spacer(1, 12 * mm))
    s.append(P(f"발행일  {gs.PUBLISH_DATE}       {gs.ENGINE_NOTE}", 10, MUTED,
               TA_CENTER, space=16))
    s.append(mktable([[Paragraph(
        f'<font name="Sans-B" color="#A93226" size="10.5">※  유의사항</font><br/>'
        f'<font name="Sans" color="#7B241C" size="9.5">{esc(gp.DISCLAIMER)}</font>',
        ParagraphStyle("d", leading=14))]], [CW],
        [("BOX", (0, 0), (-1, -1), 0.8, DISCBD),
         ("BACKGROUND", (0, 0), (-1, -1), DISCBG),
         ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
         ("LEFTPADDING", (0, 0), (-1, -1), 11), ("RIGHTPADDING", (0, 0), (-1, -1), 11)]))
    s.append(NextPageTemplate("body"))
    s.append(PageBreak())


def intro_toc(ks):
    s, d = gp.story, gp.demands
    n_rec, n_proj, n_sup = _counts(d)
    s.append(section_label("개요", before=2))
    s.append(P(f"본 보고서는 2026 MEDITEK 참여기업 {len(ks)}개사의 기술을 대상으로, MEDITEK "
               f"공급기관 {n_sup}곳이 수행한 국가 R&D 과제와의 의미 기반 매칭을 수행한 결과를 "
               f"정리한 것이다. 기업별로 적합도가 높은 추천 과제 상위 5건(총 {n_rec}건, "
               f"중복 제외 {n_proj}개 과제)을 선정하고, 매칭 근거와 상세 추천 근거를 함께 "
               f"제시하였다. {gs.METHOD_NOTE} {gs.CORPUS_NOTE}",
               10.5, INK, TA_JUSTIFY, leading=17, family="Serif", space=8))
    s.append(P("각 기업은 다음 순서로 구성된다.", 10.5, INK, family="Serif", space=3))
    for ln in ["기업 정보 — 기업명 · 수요기술(확보 희망) · 보유기술(이미 보유) · 핵심 키워드",
               f"최종 추천 과제 {gs.TOPN_LABEL} — 순위 · 과제명 · 수행기관 · 공급기관 · "
               "수행년도 · 특허 · 매칭 근거",
               "추천 과제별 상세 정보표 — 과제고유번호 · 수행기간 · 표준분류 · 연구개발단계 · "
               "수행기관 · 연구수행주체 · 연구책임자 · 국가연구자번호",
               "추천 과제별 상세 매칭 근거 — 연관성 · 기술 적합성 · 추천 과제의 우수성 · "
               "유사 사례 및 실적",
               "추천 과제별 특허 실적 — 등록·출원 구분 · 특허명 · 기관 · 국가 · 번호 · 일자"]:
        s.append(Paragraph(f'<font name="Sans-B" color="#0E7C86">· </font>'
                           f'<font name="Serif" color="#1B2430">{esc(ln)}</font>',
                           ParagraphStyle("b", fontSize=10, leading=15, leftIndent=14,
                                          spaceAfter=3)))
    s.append(P("'공급기관'은 해당 과제를 수행한 기관의 기술이전 창구(산학협력단·기술지주 등)로, "
               "기술이전 협의를 시작할 접촉 지점을 뜻한다.",
               9.5, MUTED, TA_JUSTIFY, family="Serif", leading=14, space=2))

    s.append(section_label("목차", before=14))

    def pg(v):
        return str(v) if v else "··"

    rows = [[P("번호", 9.5, HEADFG, TA_CENTER, bold=True),
             P("기업명", 9.5, HEADFG, bold=True),
             P("기술명", 9.5, HEADFG, bold=True),
             P("면", 9.5, HEADFG, TA_RIGHT, bold=True)]]
    sty = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
           ("LINEBELOW", (0, 0), (-1, -1), 0.3, HAIR),
           ("BACKGROUND", (0, 0), (-1, 0), HEADBG),
           ("TOPPADDING", (0, 0), (-1, -1), 4.5),
           ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5)]
    for k in ks:
        names = gs.tech_names(d[k])
        rows.append([P(f"No.{k}", 9, ACCENT, TA_CENTER, bold=True),
                     P(d[k]["기업명"], 9, INK, leading=11.5),
                     P("\n".join(names) if names else (d[k].get("기술분야") or ""),
                       8.4, MUTED, leading=10.6),
                     P(pg(gp.PAGE_MAP.get(k)), 9, MUTED, TA_RIGHT)])
    s.append(mktable(rows, [20 * mm, 38 * mm, CW - 70 * mm, 12 * mm], sty))
    s.append(PageBreak())


def company_list(ks):
    s, d = gp.story, gp.demands
    title = P("대상 기업 목록", 21, NAVY, bold=True, space=1)
    title._chapter_no = 1                       # 목차 페이지 산출 마커(gen_report_pdf 규약)
    s.append(title)
    s.append(mktable([[P("TARGET COMPANIES", 9, ACCENT, bold=True)]], [CW],
                     [("LINEBELOW", (0, 0), (-1, -1), 1.2, NAVY),
                      ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                      ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    s.append(Spacer(1, 3))
    s.append(P(f"기업 {len(ks)}개사", 10, MUTED, space=10))
    rows = [[P(x, 9.5, HEADFG, TA_CENTER, bold=True) for x in ("번호", "기업명", "기술명")]]
    for k in ks:
        v = d[k]
        names = gs.tech_names(v)
        rows.append([P(k, 9.5, NAVY, TA_CENTER, bold=True),
                     P(v["기업명"], 9.5, leading=12),
                     P("\n".join(names) if names else (v.get("기술분야") or ""),
                       9, INK if names else MUTED, leading=12)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)])
    for i in range(2, len(rows), 2):
        st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    s.append(mktable(rows, [16 * mm, 40 * mm, CW - 56 * mm], st))
    s.append(PageBreak())               # 요약 표는 새 페이지에서 시작(고아 표제 방지)
    supply_summary()
    s.append(PageBreak())


def supply_summary():
    s, d = gp.story, gp.demands
    s.append(section_label("공급기관별 추천 현황", before=2, after=4, size=12))
    s.append(P("추천 과제를 수행한 공급기관과 그 비중. 같은 과제가 여러 기업에 추천될 수 있어 "
               "추천 건수와 과제 수는 다를 수 있다.", 9, MUTED, space=6))
    agg = gs.supply_summary(d)
    # align 에 None 을 주면 reportlab 이 정렬 함수를 못 잡아 그리기 단계에서 죽는다(TA_LEFT 명시)
    rows = [[P(x, 9.5, HEADFG, TA_CENTER if i else TA_LEFT, bold=True)
             for i, x in enumerate(("공급기관", "추천 건수", "과제 수", "특허 건수"))]]
    for sup, a in agg:
        rows.append([P(sup, 9.5, INK, leading=12),
                     P(f"{a['rec']}건", 9.5, NAVY, TA_CENTER, bold=True),
                     P(f"{len(a['pids'])}개", 9.5, INK, TA_CENTER),
                     P(f"{a['pat']}건", 9.5, INK, TA_CENTER)])
    rows.append([P("합계", 9.5, NAVY, bold=True),
                 P(f"{sum(a['rec'] for _, a in agg)}건", 9.5, NAVY, TA_CENTER, bold=True),
                 P(f"{len(set().union(*(a['pids'] for _, a in agg)))}개", 9.5, NAVY,
                   TA_CENTER, bold=True),
                 P(f"{sum(a['pat'] for _, a in agg)}건", 9.5, NAVY, TA_CENTER, bold=True)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG),
                    ("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1), LABELBG)])
    for i in range(2, len(rows) - 1, 2):
        st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    s.append(mktable(rows, [CW - 90 * mm, 30 * mm, 30 * mm, 30 * mm], st))


def company_block(k, dm):
    s = gp.story
    badge = (f'<font name="Sans-B" color="#FFFFFF" backColor="#0E7C86"> No.{esc(k)} </font>'
             f'  <font name="Sans-B" color="#14315C" size="14">{esc(dm["기업명"])}</font>')
    from reportlab.pdfbase import pdfmetrics
    off = (pdfmetrics.stringWidth(f" No.{k} ", "Sans-B", 14)
           + pdfmetrics.stringWidth("  ", "Sans-B", 14))
    h2p = Paragraph(badge, ParagraphStyle("h2", fontName="Sans-B", fontSize=14, leading=20,
                                          spaceAfter=5, leftIndent=off,
                                          firstLineIndent=-off))
    h2p._toc_k = k                                # 목차 페이지 산출 마커
    s.append(h2p)
    s.append(mktable([[""]], [CW], [("LINEBELOW", (0, 0), (-1, -1), 0.6, HAIR)]))
    s.append(Spacer(1, 4))

    def lab(t):
        return P(t, 9.5, NAVY, TA_CENTER, bold=True)

    def prose(t):
        return P(t, 9, INK, TA_JUSTIFY, family="Serif", leading=13)

    # 수요기술(확보 희망)과 보유기술(이미 보유)을 구분 머리행으로 나눈다(docx 판과 동일)
    rows = [[lab("기업명"), P(dm["기업명"], 9)]]
    groups = []                                  # 머리행 위치(셀 합치기·음영용)
    ind = gs.industry_row(dm)
    if ind:
        rows.append([lab(ind[0]), P(ind[1], 9, INK, leading=12)])
    nm = {g: (n, gb) for g, n, gb in gs.tech_name_rows(dm)}
    body = gs.tech_rows(dm)
    for gubun in ("수요기술", "보유기술"):
        has_body = [x for x in body if x[0] == gubun]
        if gubun not in nm and not has_body:
            continue
        groups.append(len(rows))
        rows.append([P(gs.GUBUN_LABEL[gubun], 9.5, HEADFG, bold=True), ""])
        if gubun in nm:
            n, gb = nm[gubun]
            rows.append([lab("기술명"),
                         P(n + (f"\n({gb})" if gb else ""), 9, INK, leading=12)])
        for _g, label, val, _p in has_body:
            rows.append([lab(label), prose(val)])
    if (dm.get("키워드") or "").strip():
        rows.append([lab("핵심 키워드"),
                     P(" · ".join(x for x in dm["키워드"].split(";") if x), 8.8, INK,
                       leading=12)])
    st = base_grid()
    for i in range(len(rows)):
        if i in groups:
            st += [("SPAN", (0, i), (-1, i)),
                   ("BACKGROUND", (0, i), (-1, i), HEADBG)]
        else:
            st.append(("BACKGROUND", (0, i), (0, i), LABELBG))
    s.append(mktable(rows, [30 * mm, CW - 30 * mm], st))
    s.append(Spacer(1, 6))

    s.append(section_label(f"최종 추천 과제  {gs.TOPN_LABEL}", before=4, after=5, size=12))
    # 적합도 점수는 싣지 않는다(순위로만 제시)
    heads = ("순위", "과제명", "수행기관", "공급기관", "수행년도", "특허", "매칭 근거")
    rows = [[P(x, 9, HEADFG, TA_CENTER, bold=True) for x in heads]]
    for tp in dm[TOP_KEY]:
        n = npat(tp)
        rows.append([P(str(tp["rank"]), 10.5, NAVY, TA_CENTER, bold=True),
                     P(tp["과제명"], 8.4, leading=10.6),
                     P(tp.get("수행기관", ""), 8.0, INK, TA_CENTER, leading=10.2),
                     P(tp.get("공급기관", ""), 8.0, ACCENT, TA_CENTER, bold=True,
                       leading=10.2),
                     P(gp.year_cell(tp.get("과제설명문", "")), 8.0, INK, TA_CENTER),
                     P(f"{n}건", 8.0, NAVY if n else MUTED, TA_CENTER, bold=bool(n)),
                     P(gm.rename_rnd(tp.get("판단근거", "")), 8.0, INK, TA_LEFT,
                       family="Serif", leading=10.4)])
    st = base_grid([("BACKGROUND", (0, 0), (-1, 0), HEADBG)])
    for i in range(2, len(rows), 2):
        st.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    s.append(mktable(rows, [11 * mm, 42 * mm, 22 * mm, 24 * mm, 17 * mm, 10 * mm,
                            CW - 126 * mm], st))
    s.append(PageBreak())
    for tp in dm[TOP_KEY]:
        # 상세 페이지 상단 헤더는 기업명으로 표시한다(기술명이 길어 식별성이 떨어짐)
        gp.top_detail(gm.neutralize(tp), k, dm["기업명"])


def assemble():
    gp.story.clear()
    gp.story.append(NextPageTemplate("cover"))
    ks = sorted(gp.demands, key=gm.sort_key)
    cover()
    intro_toc(ks)
    company_list(ks)
    for k in ks:
        company_block(k, gp.demands[k])
    return list(gp.story)


if __name__ == "__main__":
    gp.assemble = assemble          # main() 이 모듈 전역에서 조회하므로 교체가 반영된다
    gp.main(OUT)
