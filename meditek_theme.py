# -*- coding: utf-8 -*-
"""2026 MEDITEK 보고서 테마 — 행사 정보 · CI 색상 · 로고.

색상은 짐작이 아니라 공식 자산에서 뽑았다. meditek.or.kr 의 로고 PNG 픽셀과 사이트
CSS 가 같은 값을 가리킨다.
    #1590A5  teal   — CI 의 '기술'(로고 위쪽 십자) · CSS 포인트 색
    #E94E54  coral  — CI 의 '의료·헬스케어'(로고 아래쪽 십자) · CSS 포인트 색
    #333333  charcoal — 사이트 본문색
CI 설명문의 '나비효과'(기술+의료가 만나 나비가 되는 형태)를 표지·표 머리에 두 색으로 옮겼다.

로고는 CI 페이지에서 받은 원본(MEDITEK_logo.zip)을 그대로 쓴다(assets/).

출처
    https://www.meditek.or.kr/kr/index.php
    https://www.meditek.or.kr/kr/company/intro.php
    https://www.meditek.or.kr/kr/company/ci.php
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
LOGO_H = os.path.join(ASSETS, "meditek_logo_h.png")      # 좌우조합형(표지용)
LOGO_SYMBOL = os.path.join(ASSETS, "meditek_symbol.png")  # 심볼만
LOGO_V = os.path.join(ASSETS, "meditek_logo_v.png")       # 상하조합형

# ---- CI 색상 -----------------------------------------------------------------
# 보고서 전체 룩앤필은 APOLLO CI(파랑)를 쓴다. MEDITEK 색은 로고 자체와 행사 맥락에만
# 남는다 — 행사 주최 측 브랜드가 아니라 보고서 생성 주체의 브랜드로 조판한다는 선택이다.
#
# APOLLO — 'APOLLO 로고.png' 픽셀에서 추출(불투명 화소만, 파랑 계열 82,651px 기준)
#   #095090  25,471px 30.8%  마크의 평면 브랜드 블루      → 주색
#   #1660A3     770px         마크 중간 톤               → 보조(소제목·태그)
#   #2D7DC5   1,005px         마크 밝은 톤(그라데이션 상단) → 강조(강조 바·배지)
#   #000111 391,304px         'POLLO' 워드마크 검정        → 본문 먹색(가독성 위해 살짝 들어올림)
# 옅은 배경·괘선은 주색을 흰색과 섞어 만든다(임의 색을 새로 만들지 않는다).
APOLLO_BLUE = "095090"      # 주색
APOLLO_BLUE_MID = "1660A3"  # 보조
APOLLO_BLUE_LT = "2D7DC5"   # 강조
APOLLO_INK = "16181D"       # 워드마크 검정 기반 본문색
APOLLO_GREY = "6E7681"      # 캡션·부가 정보(중성 회색)
APOLLO_HAIR = "BACEE0"      # 주색 + 흰색 72%
APOLLO_TINT = "E6EEF4"      # 주색 + 흰색 90% (라벨 셀)
APOLLO_PALE = "F4F7FA"      # 주색 + 흰색 95% (짝수행)

# MEDITEK CI — 로고 PNG 픽셀과 사이트 CSS 가 같은 값을 가리킨다. 지금은 로고·행사
# 맥락용으로만 남겨 둔다(테마를 되돌리려면 apply_* 의 표를 MEDITEK 값으로 바꾸면 된다).
TEAL = "1590A5"      # CI '기술'(로고 위쪽 십자)
TEAL_DK = "127C8D"
CORAL = "E94E54"     # CI '의료·헬스케어'(로고 아래쪽 십자)
CHARCOAL = "333333"  # 사이트 본문색
GREY = "888888"
HAIRLINE = "DDDDDD"
TEAL_TINT = "EAF4F6"
TEAL_PALE = "F5FAFB"

# gen_report / gen_report_pdf 의 색 상수 이름 → 실제 적용 색(APOLLO CI)
DOCX_COLORS = {
    "INK": APOLLO_INK, "NAVY": APOLLO_BLUE, "BLUE": APOLLO_BLUE_MID,
    "ACCENT": APOLLO_BLUE_LT, "MUTED": APOLLO_GREY, "HAIR": APOLLO_HAIR,
    "HEAD_BG": APOLLO_BLUE, "HEAD_FG": "FFFFFF",
    "LABEL_BG": APOLLO_TINT, "ZEBRA": APOLLO_PALE,
}
# 경고 박스(DISC_*)는 바꾸지 않는다 — 파랑 팔레트에서 붉은 경고가 유일한 주의 신호다.

# ---- 행사 정보(홈페이지 확인분) ---------------------------------------------
EVENT = {
    "명칭": "2026 MEDITEK – 의료기기/헬스케어 Open Innovation & Biz Partnering",
    "명칭_영문": "2026 MEDITEK – Medical Device/Healthcare Open Innovation & Biz Partnering",
    "일시": "2026년 9월 9일(수) ~ 11일(금)",
    "장소": "메종 글래드 제주",
    "주최": "MEDITEK 조직위원회",
    "주관": "대구경북첨단의료산업진흥재단 · 과학기술사업화진흥원 · 연구개발특구진흥재단 · "
          "원주의료기기산업진흥원 · 오송첨단의료산업진흥재단 · 한국기술지주회사협회 · "
          "한국연구소기술이전협회 · 한국대학기술이전협회",
    "후원": "특허법인 이노 · 젠스퀘어 · 위노베이션 · 경남김해강소특구 외 다수",
    "목적": "국내외 의료기기·헬스케어 산업의 지속 가능한 기술혁신과 성장을 목적으로 하는 "
          "기술사업화 오픈 이노베이션의 장으로, 혁신 기술 공개와 비즈니스 파트너링을 통해 "
          "공동연구·기술라이센싱·공동기술사업화·투자 연계를 도모한다.",
    "참가대상": "대학·병원·연구기관이 보유한 혁신 기술을 보유하거나 필요로 하는 기업, "
            "투자기관, 컨설팅기관 등",
    "홈페이지": "www.meditek.or.kr",
}
# 표지·개요에 싣는 순서
EVENT_ROWS = ("일시", "장소", "주최", "주관", "후원")


def apply_docx(gr):
    """gen_report 모듈의 색 상수를 MEDITEK CI 로 교체(호출 시점에 참조되므로 유효)."""
    for k, v in DOCX_COLORS.items():
        setattr(gr, k, v)
    return gr


def apply_pdf(gp):
    """gen_report_pdf 모듈의 색 상수를 APOLLO CI 로 교체."""
    from reportlab.lib import colors
    m = dict(DOCX_COLORS)
    m["HEADBG"] = m.pop("HEAD_BG")
    m["LABELBG"] = m.pop("LABEL_BG")
    m.pop("HEAD_FG", None)
    for k, v in m.items():
        setattr(gp, k, colors.HexColor("#" + v))
    gp.HEADFG = colors.white
    return gp


def logo_size(path, width_pt):
    """폭을 주면 원본 비율에 맞는 (폭, 높이)를 돌려준다."""
    from PIL import Image
    with Image.open(path) as im:
        w, h = im.size
    return width_pt, width_pt * h / w
