# -*- coding: utf-8 -*-
"""2026_MEDITEK_260820.xlsx 수요기업 × '공급기관 리스트' 수행 국가 R&D 과제 매칭.

기존 match_meditek_top10.py 는 전체 국가 R&D 코퍼스를 대상으로 했다. 이 스크립트는
같은 엑셀 안에 있는 '공급기관 리스트'(기술이전 공급 측 15개 기관)가 수행한 과제만을
후보로 삼고, 그중 특허 성과가 1건 이상 있는 과제만 매칭한다.

  ① 코퍼스   : 공급기관 수행과제 ∧ 제출년도>=YEAR_MIN ∧ 특허성과>=1건
               (제출년도는 보고서를 낸 연도다. 그 해에 과제가 살아 있었다는 뜻이므로
                결과적으로 'YEAR_MIN 이후 수행 중이거나 종료된 과제'가 남는다 —
                시작연도는 그보다 이전일 수 있고, 종료가 YEAR_MIN 이전인 과제는 없다.)
  ② 질의     : 기업별로 하나의 근거만 쓴다(우선순위 — 요구사항 그대로)
                 1) 수요기술('주요사업분야(수요기술)')이 있으면 그것
                 2) 없으면 기보유기술 내용
                 3) 둘 다 없으면 사업자번호로 기업 pkl 을 찾아 기업설명문·산업분류·키워드
               → 35B 키워드 추출 → pro-sroberta 임베딩
  ③ 1차선정  : 임베딩 코사인 유사도 + 과제 유망성점수의 가중합으로 후보 N건
  ④ 재랭킹   : 후보를 35B 로 0~100 채점(적합도) → 적합도 주도 가중합으로 최종 순위
  ⑤ 근거     : 표용 한 문장(판단근거) + 3섹션 상세근거를 35B 로 생성
               (텍스트는 전부 모델이 쓴다. 금지표현·유망성 차단은 explain_meditek_top10 재사용)
  ⑥ 산출     : <tag>_매칭.pkl / .xlsx (+ 보고서 입력 JSON) — 나중에 보고서 생성용

사용:
  python match_meditek_supply.py --dry-run          # LLM 없이 코퍼스·질의 구성만 점검
  python match_meditek_supply.py                    # 전체 실행
  python match_meditek_supply.py --verify-only      # 저장된 산출물만 검증
"""
import argparse
import json
import os
import pickle
import re
import time

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer

import compa_match as cm
import explain_meditek_top10 as xp
import match_meditek_top10 as mt
import meditek_patent_axis as pax
import meditek_phase_merge as pmg
import meditek_supply_orgs as so
from rematch_filtered import ALLOW, EMB_FILE, YEAR_MIN

SRC_XLSX = "2026_MEDITEK_260820.xlsx"
SH_DEMAND, SH_SUPPLY = "수요기업 리스트", "공급기관 리스트"
COMPANY_PKL = cm.COMPANY_EMB                  # 사업자번호 → 기업설명문(3순위 질의)
PATENT_PKL = "nice_patent_overview_rnd_linked_abstract_promise.pkl"

N_POOL = 1000              # 코사인 상위 풀(이 안에서 유사도 정규화 + 유망성 가중)
N_CAND = 60                # 1차선정 통과 후보 수(= LLM 재랭킹 대상)
FINAL = 5                  # 기업별 최종 추천 수
MIN_FIT = 40               # 적합도 하한(미달은 '보충'으로만 들어간다)
# 연계유형을 순위에 반영하는 2단 구조.
#   ① 적합도 채점은 후보 전체(N_CAND)에 그대로 걸어 척도를 흔들지 않는다
#   ② 그중 상위 KIND_POOL 건만 유형을 판정하고
#   ③ 유형 가중을 적용해 다시 정렬한 뒤 Top5 를 고른다
# 유형을 채점 프롬프트에 함께 넣으면 같은 건의 적합도가 85→40 으로 흔들려(실측) 못 쓴다.
# 가중은 적합도 약 5점 상당의 핸디캡이다 — 순위가 비슷할 때만 방향이 갈림길이 되고,
# 명확히 더 나은 공동연구 건이 기술도입 건에 밀리지는 않는다.
KIND_POOL = 15             # 유형 판정 대상(상위 후보 수)
W_KIND_CO = 0.95           # 공동연구 가중(기술도입 = 1.00)
# 1차선정 가중치 — 임베딩 유사도 + 과제 유망성 + 특허축
W_COS, W_PROM_PRE = 0.75, 0.25
# 재랭킹 가중치 — LLM 적합도가 순위를 주도한다
W_FIT, W_COS_RE, W_PROM_RE = 0.70, 0.15, 0.15
# 특허축(meditek_patent_axis) 가중 — 질의와 유사하고 권리가 살아 있는 특허를 **많이**
# 가진 과제를 끌어올린다. 두 단계 모두에 넣는다. 1차선정에 넣지 않으면 그런 과제가
# 애초에 후보(N_CAND)에 들어오지 못해 재랭킹이 손쓸 수 없다.
# 가중치 합은 코드에서 정규화되므로(wsum/psum) 실효 비중은
#   1차선정 유사도 0.625 · 유망성 0.208 · 특허 0.167
#   재랭킹   적합도 0.648 · 유사도 0.139 · 유망성 0.139 · 특허 0.074
# 재랭킹 가중을 0.12 로 두면 축이 적합도 10점 차이를 뒤집는다(적합도 10점 = 최종 6.25점,
# 특허축 0.65 차 = 최종 7.0점). 실측에서 인재양성·기관지원 사업이 적합도 50 → 40 짜리
# 추천을 밀어내고 3행 → 8행으로 늘었다. 유형 가중과 같은 '적합도 약 5점 상당' 수준으로
# 낮춘다. 1차선정은 누가 채점 대상이 되는지만 정하므로 강하게 두어도 무해하다.
W_PAT_PRE, W_PAT_RE = 0.20, 0.08
QUERY_MIN = 40             # 질의 본문이 이보다 짧으면 경고만 남긴다(질의는 그대로 사용)
# 추천이 한 건도 없으면 최소 1건을 보장할 공급기관. 점수순 선정만으로는 과제 수가 적은
# 기관이 통째로 빠지는데, 행사 운영상 기관별 미팅 기회를 보장해야 할 때 쓴다.
# 넣을 기관은 손으로 고른다 — 전 기관에 일괄 적용하면 최고 적합도가 40(하한)인 기관까지
# 끌려 들어와 85점 추천을 밀어낸다(국립경국대: 의료·제약 과제 7건, 최고 40점).
#   한양대 에리카: 의료·제약 과제 19건, 최고 80점((주)비전메디컬) → 손실이 작다
COVER_SUPPLIERS = ("한양대학교 에리카산학협력단",)
# 질의 근거는 35B 가 판정한다(select_bases) — 수요기술 설명이 불충분하면 기보유기술,
# 그래도 불충분하면 기업DB 사업내용을 단계적으로 더한다. 사용자가 기업 번호를 지정할
# 필요는 없고, --augment 는 판정을 무시하고 강제로 더하는 예외 수단으로만 남긴다.
DEM_MAX, HOLD_MAX = 1200, 1600     # 채점 프롬프트에 넣는 본문 절단 길이


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- 기술 분야 한정
# MEDITEK 은 의료·제약 분야 행사이므로 무관한 분야(에너지·정보보호·건설·원자력 등)의 과제가
# 추천되지 않게 코퍼스를 좁힌다. 두 갈래로 판정한다.
#   ① 과학기술표준분류(중)가 의료·제약·생명 계열이면 통과
#   ② 분류가 그 밖이어도 과제명·설명문·키워드에 의료 관련 어휘가 MIN_TERMS 종 이상 나오면 통과
#      — 의료기기·의료영상 과제가 '마이크로기계시스템'·'고분자재료'·'계측기기'로 분류되는
#        일이 많아 ①만으로는 정작 필요한 과제가 빠진다.
#   단, ②로도 살릴 수 없는 분야(농림수산·에너지·건설·원자력·국방·교통)는 따로 막는다.
MED_FIELDS = {
    # 의학·보건·의료기기
    "임상의학", "의생명과학", "진단기기", "의료기기", "의료정보 및 시스템", "기타 보건의료",
    "보조 및 복지기기", "치의과학", "한의과학", "수의과학", "면역학 및 생리학", "생리학",
    "뇌공학", "뇌의약", "뇌인지", "보건학", "보건의료 생명", "치의학 생명", "의치과학",
    "간호과학", "방사선기술", "환경보건", "치료 및 진단기기", "의료기기안전관리", "영양관리",
    # 건강기능식품 — MEDITEK 참여기업에 건강기능식품 수요가 있어 식품 계열을 포함한다
    "식품과학", "식품영양과학",
    # 제약·의약품
    "의약품개발기술", "신약·의약품개발", "의약품 및 의약품개발기술",
    "의약품 및 의약품 개발 기술", "의약품안전관리", "독성 및 안전성관리 기반기술",
    "독성·안전성 평가·관리", "위해성 평가 및 관리",
    # 생명과학·바이오(의료·제약 직결)
    "분자세포생물학", "융합바이오", "생물공학", "신경생물학", "유전공학",
    "유전학 및 유전공학", "유전학·유전체학", "생화학 및 구조생물학", "생화학", "구조생물학",
    "기타생명과학", "기타 생명과학", "산업바이오", "생물위해성", "생물화학 공정기술",
}
NON_MED_FIELDS = {
    "작물보호", "원예특용작물과학", "식량작물과학", "농생물학", "농화학", "농업환경생태",
    "농업인프라 공학", "농업·식품 기계·설비", "농업기계 및 설비", "동물자원과학",
    "산림자원학", "임산공학", "수산양식", "해양자원", "기타 농림수산식품",
    "농수축산물 안전", "농수축산물 품질·안전 관리",
    "신재생에너지", "전지", "수화력발전", "가스에너지", "스마트그리드", "배전계통",
    "원자력안전기술", "원자력소재", "원자로 노심 기술", "기타 원자력", "핵융합", "화력탄약",
    "건설시공 및 재료", "건설 환경설비 기술", "시설물 안전 및 유지관리 기술",
    "시설물 설계 및 해석기술", "국토공간개발기술", "국토정책 및 계획", "지역개발",
    "물관리", "수공시스템기술", "유지관리 기술",
    "철도차량", "철도교통기술", "도로교통기술", "교통", "물류기술", "항공시스템",
    "우주발사체", "우주시스템", "인공위성",
    "국방소재", "국방정보통신", "국방플랫폼",
    "폐기물 관리 및 자원순환", "자원순환", "온실가스 처리", "대기질 관리", "대기과학",
    "기후학", "기후과학", "기상과학", "지질과학", "해양과학", "해양시스템",
    "기타 지구과학", "자원",
}
MED_TERMS = re.compile(
    r"의료|의학|임상|진단|치료|환자|병원|질환|질병|약물|의약|제약|신약|백신|항체|항균|"
    r"바이오|생체|생물|세포|단백질|유전자|유전체|효소|면역|미생물|박테리아|바이러스|"
    r"암\b|종양|뇌|신경|심장|혈액|혈관|호흡|폐\b|간\b|신장|골\b|뼈|관절|치아|치과|구강|"
    r"피부|안구|망막|재활|헬스케어|건강|수술|내시경|카테터|스텐트|임플란트|보청기|초음파|"
    r"MRI|CT\b|PET\b|X-ray|PCR|오가노이드|줄기세포|생검|채혈|투석|봉합|마취|약제|처방|"
    r"병리|해부|호르몬|대사|염증|감염")
MIN_TERMS = 4              # ② 통과에 필요한 의료 어휘 종류 수
# ③ 건강기능식품 전용 어휘 — 한 개만 나와도 통과시킨다. 건강기능식품 소재 연구는
#    '산림자원학'·'원예특용작물과학' 처럼 농림 계열로 분류되는 일이 많아 ①②로는 못 살린다.
#    '추출물'·'영양' 같은 넓은 낱말은 넣지 않는다(화학·농업 과제가 통째로 끌려온다).
HFF_TERMS = re.compile(
    r"건강기능식품|기능성\s?식품|고령친화식품|건강식품|식품소재|식품\s?소재|"
    r"프로바이오틱스|프리바이오틱스|유산균|식이섬유|기능성\s?소재|건강기능성|"
    r"영양성분|식이보충|영양보충|메디푸드|특수의료용도식품")


def med_terms_count(*texts):
    """의료 관련 어휘가 몇 '종' 나오는지(같은 낱말 반복은 1종)."""
    return len(set(MED_TERMS.findall(" ".join(str(t) for t in texts if t))))


def in_medical_domain(field, name, desc):
    """의료·제약·건강기능식품 분야 연관 여부 → (통과?, 근거).

    판정 입력은 과제명·과제설명문뿐이다. 둘 다 산출물(pkl/xlsx)에 그대로 실리므로
    코퍼스를 다시 만들지 않고도 검증이 같은 결과를 재현할 수 있다.
    """
    if field in MED_FIELDS:
        return True, f"분류 '{field}'"
    hff = HFF_TERMS.search(f"{name} {desc}")
    if hff:
        return True, f"건강기능식품 어휘 '{hff.group(0)}'"
    n = med_terms_count(name, desc)
    if field in NON_MED_FIELDS:
        return False, f"분류 '{field}'(비의료 분야, 어휘 {n}종)"
    if n >= MIN_TERMS:
        return True, f"의료 어휘 {n}종"
    return False, f"분류 '{field or '미상'}' · 어휘 {n}종"


# ---------------------------------------------------------------- 공급기관 → 과제수행기관명
# 대응표는 meditek_supply_orgs.SUPPLY_ORGS 에 있다(코퍼스의 고유 과제수행기관명 57,669개를
# 전수 조회해 확정한 기관명 목록 + 범위 메모). 정규식으로 그때그때 짐작하지 않는다 —
# '울산과학대학교'(UNIST 아님)·'서울아산병원'(울산대 산하 아님) 처럼 이름만 비슷한 기관이
# 섞이거나, 반대로 'AI' 표기 변형을 놓치는 일을 막기 위한 것이다.
# 범위: 대학 산학협력단·기술지주 = 그 대학 산하 전부 / 고려대 의료원 = 병원 계열만 /
#       한양대 에리카 = ERICA 캠퍼스만 / 백병원·아산병원 계열 제외.
def resolve_supply(sup_names, org_names):
    """공급기관명 → 코퍼스에 실제로 있는 과제수행기관명 목록.

    반환 (해석표, 대응표에 없는 공급기관, 대응표에는 있으나 코퍼스에 없는 기관명).
    """
    resolved, unknown, missing = {}, [], {}
    have = set(org_names)
    for name in sup_names:
        ent = so.SUPPLY_ORGS.get(name)
        if ent is None:
            unknown.append(name)
            continue
        memo, _probe, orgs = ent
        hit = sorted(o for o in orgs if o in have)
        gone = [o for o in orgs if o not in have]
        if gone:
            missing[name] = gone
        resolved[name] = (hit, memo)
    return resolved, unknown, missing


# ---------------------------------------------------------------- 매칭 단위(수요기업)
def _s(v):
    """엑셀 셀 → 정리된 문자열(결측·자리표시자는 '')."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = re.sub(r"_x00[0-9A-Fa-f]{2}_|[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(v)).strip()
    return "" if s.lower() in ("nan", "none", "-", "") else s


def _brn(v):
    """사업자등록번호 → 숫자만(기업 pkl 의 사업자번호와 같은 형식)."""
    return re.sub(r"\D", "", _s(v))


def company_index(src=COMPANY_PKL):
    """기업 pkl → {사업자번호(숫자만): 기업 레코드}. 설명문이 긴 행을 채택."""
    cdf = pd.read_pickle(src)
    keep = ["사업자번호", "한글업체명", "기업설명문", "10차산업코드명", "키워드_리스트",
            "desc_ok", "desc_issue"]
    cdf = cdf[[c for c in keep if c in cdf.columns]]
    idx = {}
    for r in cdf.to_dict("records"):
        k = _brn(r.get("사업자번호"))
        if not k:
            continue
        if k in idx and len(_s(idx[k].get("기업설명문"))) >= len(_s(r.get("기업설명문"))):
            continue
        idx[k] = r
    del cdf
    return idx


def _hold_from_sheet(r):
    """260820 시트의 기보유기술 컬럼 → (기술명, 합본 본문, 기술유형, 기술분야)."""
    return (_s(r.get("기보유기술명")), _s(r.get("기보유기술 내용")),
            _s(r.get("기술유형")), _s(r.get("기술분야")))


def _hold_from_company(rec):
    """기업 pkl 레코드 → 기보유 슬롯에 넣을 사업내용 합본(라벨 보존)."""
    desc, ind = _s(rec.get("기업설명문")), _s(rec.get("10차산업코드명"))
    kws = rec.get("키워드_리스트")
    kws = [str(x) for x in kws][:20] if isinstance(kws, (list, tuple, np.ndarray)) else []
    parts = []
    if desc:
        parts.append(f"[사업내용]\n{desc}")
    if ind:
        parts.append(f"[산업분류]\n{ind}")
    if kws:
        parts.append("[기업 키워드]\n" + ", ".join(kws))
    return "\n\n".join(parts), ind


def load_units(xlsx=SRC_XLSX, cidx=None):
    """수요기업 리스트 → 매칭 단위. 여기서는 '쓸 수 있는 근거'까지만 모은다.

    실제로 어떤 근거를 질의에 쓸지는 select_bases() 가 35B 판정으로 정한다(모델 로드 이후).
    같은 기업이 여러 행에 있으면 한 단위로 합치고 수요기술 본문은 서로 다른 내용만 이어붙인다.
    세 근거가 모두 없으면 질의를 만들 수 없어 제외한다(사업자번호가 기업 pkl 에 없는 경우 —
    기업명으로 유추하지 않는다).
    """
    df = pd.read_excel(xlsx, sheet_name=SH_DEMAND)
    merged = {}
    for r in df.to_dict("records"):
        name = _s(r.get("기관명"))
        if not name:
            continue
        key = cm.norm_name(name)
        u = merged.setdefault(key, {
            "키": key, "기업명": name, "번호": _s(r.get("No")),
            "사업자등록번호": _s(r.get("사업자등록번호")), "대표자": _s(r.get("대표자")),
            "기관유형": _s(r.get("기관유형")), "년도": _s(r.get("년도")),
            "행수": 0, "수요_본문들": [],
        })
        u["행수"] += 1
        dem = _s(r.get("주요사업분야(수요기술)"))
        if dem and not any(dem in x or x in dem for x in u["수요_본문들"]):
            u["수요_본문들"].append(dem)
        hname, hbody, htype, hfield = _hold_from_sheet(r)
        if hbody and not u.get("_hold"):
            u["_hold"] = (hname, hbody, htype, hfield)
            u["관리번호"] = _s(r.get("관리번호"))

    keep, skip = [], []
    for u in merged.values():
        dem = "\n".join(u.pop("수요_본문들"))
        hname, hbody, htype, hfield = u.pop("_hold", ("", "", "", ""))
        u.setdefault("관리번호", "")
        u["기보유기술_보유"] = bool(hbody)
        brn = _brn(u["사업자등록번호"])
        crec = (cidx or {}).get(brn)
        u["기업DB_기업명"] = _s(crec.get("한글업체명")) if crec else ""
        cbody, cind = _hold_from_company(crec) if crec else ("", "")

        # 우선순위대로 '쓸 수 있는' 근거를 모아 둔다 — 선택은 select_bases() 에서
        avail = {}
        if dem:
            avail["수요기술"] = {"명": _first_line(dem), "본문": dem, "유형": "", "분야": ""}
        if hbody:
            avail["기보유기술"] = {"명": hname, "본문": hbody, "유형": htype, "분야": hfield}
        if cbody:
            avail["기업DB"] = {"명": "", "본문": cbody, "유형": "", "분야": cind}
        if not avail:
            u["소스"] = "없음"
            why = ("기업 pkl 에 설명문·키워드 없음" if crec
                   else f"수요·기보유 없음 + 사업자번호 {brn} 기업 pkl 미등록")
            skip.append(dict(u, 제외사유=why, 근거=[], 보조사용=False, 질의길이=0,
                             수요기술명="", **{"수요기술 내용": "", "기보유기술 내용": ""},
                             기보유기술명="", 기술유형="", 기술분야=""))
            continue
        u["_avail"] = avail
        apply_bases(u, [next(b for b in BASIS_ORDER if b in avail)])
        keep.append(u)
    keep.sort(key=lambda u: float(u["번호"]) if u["번호"].replace(".", "").isdigit() else 1e9)
    return keep, skip


BASIS_ORDER = ("수요기술", "기보유기술", "기업DB")


def apply_bases(u, bases):
    """선택된 근거 목록을 단위의 질의 필드에 반영한다.

    기보유기술과 기업DB 는 프롬프트·보고서에서 같은 자리('기보유기술 내용')를 쓰므로,
    둘이 함께 선택되면 라벨을 유지한 채 이어붙인다(양쪽 모두 '[라벨]\n본문' 형식이다).
    """
    av = u["_avail"]
    bases = [b for b in BASIS_ORDER if b in bases and b in av]
    u["근거"] = bases
    u["소스"] = "+".join(bases)
    u["보조사용"] = len(bases) > 1
    d = av.get("수요기술") if "수요기술" in bases else None
    u["수요기술명"] = d["명"] if d else ""
    u["수요기술 내용"] = d["본문"] if d else ""
    hold = [av[b] for b in ("기보유기술", "기업DB") if b in bases]
    u["기보유기술명"] = next((h["명"] for h in hold if h["명"]), "")
    u["기보유기술 내용"] = "\n\n".join(h["본문"] for h in hold)
    u["기술유형"] = next((h["유형"] for h in hold if h["유형"]), "")
    u["기술분야"] = next((h["분야"] for h in hold if h["분야"]), "")
    u["질의길이"] = len(u["수요기술 내용"]) + len(u["기보유기술 내용"])
    return u


# ---------------------------------------------------------------- 근거 충분성 판정(35B)
_JUDGE_SYS = (
    "당신은 기업의 기술 정보를 읽고 국가 R&D 과제를 검색해 주는 기술이전 전문가입니다. "
    "주어진 설명만으로 '어떤 기술 분야의 과제를 찾아야 하는지' 특정할 수 있는지 판정합니다.")
_JUDGE_GUIDE = (
    "아래 [기업 기술 설명] 만 보고, 연계할 국가 R&D 과제를 찾기에 충분한지 판정하세요.\n"
    "  충분: 구체적인 기술·제품·응용 분야가 드러나 검색 키워드를 뽑을 수 있다.\n"
    "  불충분: 업종명·사업분야명 나열이거나 한두 낱말·관심 표명 수준이어서 어떤 기술을 "
    "찾아야 할지 특정되지 않는다.\n"
    "분량이 짧아도 기술 내용이 특정되면 충분입니다. 길어도 업종 나열뿐이면 불충분입니다.\n"
    '반드시 아래 JSON 형식으로만 답하세요(이유는 40자 이내 한국어).\n'
    '{"충분": true 또는 false, "이유": "<근거>"}')
JUDGE_MAX = 1800            # 판정 프롬프트에 넣는 본문 절단 길이


def _parse_judge(text):
    m = re.search(r'\{[^{}]*"충분"[^{}]*\}', text, re.DOTALL)
    if not m:
        return None, ""
    try:
        d = json.loads(m.group(0))
    except Exception:
        return None, ""
    v = d.get("충분")
    if isinstance(v, str):
        v = v.strip().lower() in ("true", "yes", "y", "충분", "1")
    if not isinstance(v, bool):
        return None, ""
    return v, str(d.get("이유", "")).strip()[:80]


def judge_enough(u, ckpt=None, use_llm=True):
    """현재 선택된 근거만으로 과제 검색이 가능한지 35B 판정 → (충분?, 이유, 출처).

    호출 실패·파싱 실패 시에는 조용히 통과시키지 않고 길이 규칙(QUERY_MIN)으로 대체한다.
    """
    ck = f"{u['키']}::{u['소스']}"
    if ckpt is not None and ck in ckpt:
        ok, why = ckpt[ck]
        return bool(ok), why, "ckpt"
    if not use_llm:                                 # dry-run: 모델 없이 길이 규칙
        ok = u["질의길이"] >= QUERY_MIN
        return ok, f"길이 규칙({u['질의길이']}자)", "길이"
    doc = query_doc(u)[:JUDGE_MAX]
    msgs = [{"role": "system", "content": _JUDGE_SYS},
            {"role": "user", "content": f"[기업] {u['기업명']}\n\n[기업 기술 설명]\n{doc}"
                                        f"\n\n{_JUDGE_GUIDE}"}]
    try:
        ok, why = _parse_judge(cm.stream_explanation(msgs, max_tokens=200, temperature=0.0,
                                                    top_p=1.0))
    except Exception as e:
        ok, why = None, f"판정 호출 실패: {e}"
    src = "35B"
    if ok is None:                                  # 폴백 — 길이 규칙
        ok = u["질의길이"] >= QUERY_MIN
        why = (why or "판정 응답 없음") + f" → 길이 규칙 대체({u['질의길이']}자)"
        src = "폴백"
    if ckpt is not None:
        ckpt[ck] = [bool(ok), why]
    return bool(ok), why, src


def select_bases(units, ckpt=None, force=(), disable=False, log_fn=None, use_llm=True):
    """기업별 질의 근거를 단계적으로 확정한다.

    우선순위 첫 근거로 시작해, 35B 가 '불충분'이라고 보면 다음 순위 근거를 더한다
    (수요기술 → +기보유기술 → +기업DB). 더할 근거가 없으면 그대로 둔다.
    force 에 든 기업 No/키는 판정 없이 다음 근거를 더한다(수동 지정).
    """
    out = []
    for u in units:
        avail = [b for b in BASIS_ORDER if b in u["_avail"]]
        forced = u["번호"] in set(force) or u["키"] in set(force)
        trace = []
        while True:
            nxt = next((b for b in avail if b not in u["근거"]), None)
            if nxt is None:
                break
            if disable and not forced:
                break
            if forced:
                trace.append((u["소스"], True, "수동 지정", "force"))
                apply_bases(u, u["근거"] + [nxt])
                continue
            ok, why, src = judge_enough(u, ckpt, use_llm)
            trace.append((u["소스"], ok, why, src))
            if ok:
                break
            apply_bases(u, u["근거"] + [nxt])
        u["근거판정"] = trace
        out.append(u)
        if log_fn:
            for src_label, ok, why, how in trace:
                log_fn(f"      [{u['번호']:>3}] {u['기업명']:<16} {src_label:<18}"
                       f"{'충분' if ok else '불충분':<5}({how}) {why}")
            if len(u["근거"]) > 1:
                log_fn(f"      [{u['번호']:>3}] {u['기업명']:<16} → 최종 근거 {u['소스']} "
                       f"(질의 {u['질의길이']}자)")
    return out


# 본문 첫 줄이 라벨·머리말뿐인 경우(제출 서식의 항목명) — 기술명으로 쓰지 않고 다음 줄로 넘어간다
_LABEL_ONLY = {"기술명", "수요기술명", "기술개요", "기술 개요", "수요기술개요", "수요기술 개요",
               "기업개요", "기업 개요", "개요", "회사소개", "회사 소개", "주요기능", "주요 기능",
               "수요기술내용", "수요기술 내용", "제품개요", "제품 개요"}
# 목록 머리표·번호(1. / 1) / 가. / - / • / ※) 와 통째로 감싼 대괄호
_LEAD = re.compile(r"^\s*(?:\d+\s*[.)]|[가-하]\s*[.)]|[-*•·※○□▪]+)\s*")
_WRAP = re.compile(r"^\[(.+)\]$", re.DOTALL)
# 문장 끝(마침표+공백/줄끝, 물음표·느낌표·。). '·' 는 '뷰티·헬스케어' 처럼 낱말 안에 쓰이므로 제외한다
_SENT_END = re.compile(r"(?:\.\s|\.$|[。?!])")


def _first_line(text, maxlen=60):
    """수요기술명 컬럼이 없으므로 본문에서 표시용 기술명을 뽑는다(최대 maxlen 자).

    제출 원문이 제각각이라 그대로 첫 조각을 쓰면 '1'(번호만) · '기술명'(항목 라벨) ·
    '- SUS 코일…'(머리표 잔재) 같은 값이 나온다. 줄 단위로 훑어 머리표·번호·대괄호를
    떼고, 라벨만 있는 줄은 건너뛴 뒤 첫 문장을 취한다.
    """
    for raw in str(text or "").split("\n"):
        line = _LEAD.sub("", raw.strip()).strip()
        m = _WRAP.match(line)
        if m:
            line = m.group(1).strip()
        if not line:
            continue
        if re.sub(r"\s+", "", line) in {re.sub(r"\s+", "", x) for x in _LABEL_ONLY}:
            continue                            # '기술명' 처럼 항목 라벨뿐인 줄
        cut = _SENT_END.search(line)
        head = (line[:cut.start()] if cut else line).strip().rstrip(":：")
        if len(head) < 4:                       # 번호·기호만 남은 조각
            continue
        return (head[:maxlen] + "…") if len(head) > maxlen + 1 else head
    return ""


def query_doc(u):
    """키워드 추출용 문서 — u['근거'] 에 있는 근거만 순서대로 싣는다."""
    parts = []
    for b in u["근거"]:
        if b == "수요기술":
            parts.append(f"[확보하려는 수요기술] {u['수요기술명']}\n{u['수요기술 내용']}")
        elif b == "기보유기술":
            parts.append(f"[이미 보유한 기술·사업내용] {u['기보유기술명']}\n"
                         f"{u['기보유기술 내용']}")
        else:
            parts.append(f"[기업 사업내용]\n{u['기보유기술 내용']}")
    return "\n\n".join(parts)


def keywords_of(u):
    return mt.filter_keywords(cm.extract_keywords_doc(query_doc(u)))


# ---------------------------------------------------------------- 코퍼스
# 코퍼스에 넣을 특허의 상태 — 출원 계류(공개) 또는 등록만 인정한다.
# 거절·취하·포기는 권리가 성립하지 않았고 소멸은 권리가 끝났다. 그런 특허만 가진
# 과제는 기술이전 협의 대상이 될 수 없으므로 코퍼스에서 뺀다.
#   등록 1.00 / 공개 0.45  → 인정      거절 0.10 / 취하·포기 0.05 / 소멸 0.55 → 배제
# (소멸의 법적상태_score 가 0.55 로 공개보다 높지만 권리는 이미 끝났다. 점수가 아니라
#  상태명으로 판정해야 하는 이유다.)
ALIVE_STATES = ("등록", "공개")


def patent_counts(src=PATENT_PKL, alive_only=True):
    """nice_patent pkl → {과제고유번호: 특허 건수}. 보고서 특허 표와 같은 소스.

    alive_only=True 면 상태가 ALIVE_STATES 인 특허만 센다.
    """
    df = pd.read_pickle(src)
    if alive_only:
        col = "특허등록상태명"
        if col not in df.columns:
            raise SystemExit(f"[중단] {src} 에 {col} 컬럼이 없어 상태 판정 불가")
        n0 = len(df)
        df = df[df[col].astype(str).str.strip().isin(ALIVE_STATES)]
        log(f"· 특허 상태 필터: {n0} → {len(df)}건 "
            f"(인정 {'/'.join(ALIVE_STATES)} · 배제 거절·취하·포기·소멸)")
    cnt = {}
    for v in df["과제번호"].values:
        pids = (v if isinstance(v, (list, tuple, set, np.ndarray))
                else re.findall(r"\d{6,}", str(v)))
        for p in pids:
            p = str(p).strip()
            if p:
                cnt[p] = cnt.get(p, 0) + 1
    del df
    return cnt


def patent_sets(src=PATENT_PKL, alive_only=True):
    """{과제고유번호: {특허출원번호…}} — 연차 통합 시 합집합을 내려면 건수가 아니라
    출원번호 집합이 필요하다(연차마다 같은 특허가 다시 실린다)."""
    df = pd.read_pickle(src)
    if alive_only:
        df = df[df["특허등록상태명"].astype(str).str.strip().isin(ALIVE_STATES)]
    out = {}
    for v, no in zip(df["과제번호"].values, df["특허출원번호"].values):
        pids = (v if isinstance(v, (list, tuple, set, np.ndarray))
                else re.findall(r"\d{6,}", str(v)))
        k = str(no).strip()
        for q in pids:
            q = str(q).strip()
            if q:
                out.setdefault(q, set()).add(k)
    del df
    return out


def promising_patent_sets(min_score=80.0, states=("등록",), src=PATENT_PKL):
    """{과제고유번호: {유망 등록특허 출원번호…}} — 연차 통합 시 합집합용."""
    df = pd.read_pickle(src)
    df = df[df["특허등록상태명"].astype(str).str.strip().isin(states)]
    df = df[pd.to_numeric(df["유망성점수"], errors="coerce") >= min_score]
    out = {}
    for v, no in zip(df["과제번호"].values, df["특허출원번호"].values):
        pids = (v if isinstance(v, (list, tuple, set, np.ndarray))
                else re.findall(r"\d{6,}", str(v)))
        k = str(no).strip()
        for q in pids:
            q = str(q).strip()
            if q:
                out.setdefault(q, set()).add(k)
    del df
    return out


def promising_patent_counts(min_score=80.0, states=("등록",), src=PATENT_PKL):
    """{과제고유번호: 유망 등록특허 건수}.

    '유망 특허' = 상태가 등록이고 특허 유망성점수가 min_score 이상.
    등록 특허의 유망성점수 분포는 코퍼스 실측 P25 82.3 · P50 88.5 이므로 80은
    등록 특허의 하위 4분위를 걸러내는 선이다.
    이 값은 **과제명 예외 판정에만** 쓰고 매칭 점수에는 넣지 않는다. 점수에 넣으면
    합성 유망성점수 안의 모과제유망성이 과제 유망성 축과 이중 계산된다.
    """
    df = pd.read_pickle(src)
    need = {"과제번호", "특허등록상태명", "유망성점수"}
    if not need <= set(df.columns):
        raise SystemExit(f"[중단] {src} 에 없는 컬럼: {sorted(need - set(df.columns))}")
    df = df[df["특허등록상태명"].astype(str).str.strip().isin(states)]
    df = df[pd.to_numeric(df["유망성점수"], errors="coerce") >= min_score]
    cnt = {}
    for v in df["과제번호"].values:
        pids = (v if isinstance(v, (list, tuple, set, np.ndarray))
                else re.findall(r"\d{6,}", str(v)))
        for q in pids:
            q = str(q).strip()
            if q:
                cnt[q] = cnt.get(q, 0) + 1
    del df
    return cnt


# ---- 과제명 정상성 -----------------------------------------------------------
# 과제명이 연구 주제가 아니라 기관·조직 이름이거나 인력양성·기관지원 사업 이름인
# 과제가 코퍼스에 섞여 있다(예: '충북대학교', '생태환경독성연구소',
# '지역지능화혁신인재양성'). 이런 과제는 이전받을 기술이 특정되지 않아 기술이전
# 보고서의 추천 대상이 될 수 없다. 게다가 산하 연구실 특허를 대량 누적해(최대 153건)
# 특허축에 유리해서 정상 과제를 밀어낸다.
#   ① 기관·조직명   과제명이 기관 단위 명사로 끝난다. '연구실'은 제외한다 —
#                   '간-대장 상통 대사항상성 연구실'처럼 주제를 담은 정상 과제명이다.
#   ② 인력양성 사업  인재양성·대학원지원·창업선도 등 사업 이름
#   ③ 주제 없는 단문 12자 이하이면서 연구 주제 어휘가 하나도 없다
#                   ('차세대 창의약학' 배제 / '암화 방어기전 연구'는 '연구'가 있어 유지)
# 실측: 코퍼스 1,731건 중 46건(2.7%)이 걸리고 오탐은 없었다.
NAME_ORG_TAIL = re.compile(
    r"(대학교|대학|산학협력단|기술지주|의료원|병원|연구원|연구소|센터|재단|공단|공사"
    r"|사업단|본부)$")
NAME_PROGRAM = re.compile(
    r"인재양성|인력양성|대학원지원|창업선도|혁신인재|리더스|육성사업|지원사업"
    r"|양성사업|장학|교육과정")
NAME_SUBJ = re.compile(
    r"개발|연구|규명|구축|평가|분석|진단|치료|합성|설계|제작|검증|응용|기반|이용"
    r"|활용|시스템|플랫폼|기술|공정|소재|모델|탐색|발굴|측정|제어|예측")
NAME_MIN = 12              # 이 길이 이하에서만 '주제 어휘 없음'을 배제 사유로 본다
# 예외 조항 — 과제명이 연구 주제가 아니어도 아래 둘을 **모두** 충족하면 남긴다.
# 사업 단위 묶음이라도 실제로 이전할 만한 기술이 확인되는 경우를 버리지 않기 위한 것이다.
#   ① LLM 적합도 >= NAME_EXEMPT_FIT
#   ② 등록 상태이고 유망성점수 >= NAME_PROMISE_MIN 인 특허를 1건 이상 보유
# 예외 판정에 적합도가 필요해서 과제명 검사는 코퍼스에서 배제하지 않고 표시만 하고,
# LLM 채점이 끝난 뒤 최종 선정 단계에서 걸러낸다.
# 실측(코퍼스 1,731건 중 과제명 비정상 46건): 적합도 80 이상 6건 → 그중 유망 등록특허를
# 가진 5건이 남고, 등록특허가 없는 '2024년도 대학기술경영촉진사업_충북대학교'만 빠진다.
# 예외를 절대 인정하지 않는 과제명 — **기관 이름 단독**('충북대학교', '경북대학교').
# 이름에 연구 주제가 한 조각도 없어서, 적합도가 높아도 무엇을 이전받는지 특정할 수 없다.
# '연구소'·'센터'로 끝나는 이름은 제외한다 — '약학기술연구소', '나노센서 연구소'처럼
# 연구 영역을 담고 있어 상담 단서가 된다. 공백 없는 한 낱말만 본다.
NAME_ORG_ONLY = re.compile(
    r"^[가-힣A-Za-z0-9()]+(대학교|대학|연구원|의료원|병원|산학협력단|기술지주|재단"
    r"|공단|공사)$")
NAME_EXEMPT_FIT = 80       # 예외 인정 적합도 하한
NAME_PROMISE_MIN = 80      # '유망 특허' 판정 기준(등록 특허 유망성점수 실측 P25 = 82.3)
NAME_EXTRA = 15            # 과제명 비정상 과제를 후보 정원(N_CAND) 밖에서 추가 채점할 수
# 근거문에 인용할 특허 수 — 질의 유사도 P90 을 넘는 특허만, 유사도 상위 이 개수까지
# 프롬프트에 넣는다. 넘는 특허가 없으면 빈 목록이 되고 프롬프트 규칙상 언급하지 않는다.
CITE_PATENTS = 3


def org_only_name(name):
    """과제명이 기관 이름 단독인가 — 이 경우 예외(적합도·유망특허)를 인정하지 않는다."""
    return bool(NAME_ORG_ONLY.match(re.sub(r"\s+", "", str(name)).strip()))


def bad_project_name(name):
    """과제명이 연구 주제가 아닌 경우 → (True, 사유). 정상이면 (False, "")."""
    t = re.sub(r"\s+", " ", str(name)).strip()
    if org_only_name(t):
        return True, "기관명 단독"
    if NAME_ORG_TAIL.search(t):
        return True, "기관·조직명"
    if NAME_PROGRAM.search(t):
        return True, "인력양성·기관지원 사업"
    if len(t) <= NAME_MIN and not NAME_SUBJ.search(t):
        return True, "연구주제 없는 단문"
    return False, ""


def build_corpus(sup_names, year_min=YEAR_MIN, patent_src="nice", allow_subject=False,
                 field="의료제약", patent_status="생존", check_name=True,
                 promise_min=NAME_PROMISE_MIN, merge_phase=True):
    """공급기관 수행 ∧ 최근 제출 ∧ **생존 특허**(출원 계류 또는 등록) 1건 이상인 코퍼스."""
    t0 = time.time()
    log(f"· 과제 임베딩 로드… ({EMB_FILE})")
    pdf = pd.read_pickle(EMB_FILE)
    pid_all = pdf["과제고유번호"].astype(str).values
    log(f"  전체 {len(pdf)}건 ({time.time()-t0:.0f}s)")

    log("· 과제 메타(수행기관·특허·논문) 로드…")
    with open(cm.PROJECT_META, "rb") as fp:
        pmeta = pickle.load(fp)
    org_all = np.array([str(pmeta.get(p, {}).get("과제수행기관명", "")) for p in pid_all])

    resolved, unknown, missing = resolve_supply(sup_names, set(org_all))
    if unknown:
        raise SystemExit(f"[중단] meditek_supply_orgs.SUPPLY_ORGS 에 없는 공급기관: {unknown}")
    if missing:
        raise SystemExit(f"[중단] 대응표에 있으나 코퍼스에 없는 과제수행기관명: {missing} "
                         "— 코퍼스가 바뀌었는지 확인하세요")
    _miss, unclassified = so.audit(set(org_all))
    n_unc = sum(len(v) for v in unclassified.values())
    log(f"  대응표 감사: 미분류(제외 검토 완료) 기관명 {n_unc}개 "
        "— export_supply_orgs.py 의 '제외검토' 시트 참조")
    log(f"· 공급기관 {len(resolved)}곳 → 과제수행기관명 해석")
    ok_orgs = set()
    resolved_rev = {}                 # 과제수행기관명 → 공급기관(연차 통합 키에 쓴다)
    for name, (hit, memo) in resolved.items():
        ok_orgs.update(hit)
        for h in hit:
            resolved_rev.setdefault(h, name)
        log(f"    {name:<22} 기관 {len(hit):>2}개  [{memo}]"
            + (f" ← {', '.join(hit[:5])}{' …' if len(hit) > 5 else ''}" if hit
               else "  ← 코퍼스에 수행과제 없음"))

    y = pd.to_numeric(pdf["제출년도"], errors="coerce").values
    m_org = np.isin(org_all, sorted(ok_orgs))
    m_year = y >= year_min
    pat_meta = np.array([mt._to_int(pmeta.get(p, {}).get("특허건수")) for p in pid_all],
                        dtype=np.int32)
    log(f"· 특허 성과 연계 로드… ({PATENT_PKL})")
    pcnt = patent_counts(alive_only=(patent_status == "생존"))
    pat_nice = np.array([pcnt.get(p, 0) for p in pid_all], dtype=np.int32)

    m = m_org & m_year
    m_nice, m_pm = pat_nice >= 1, pat_meta >= 1
    log(f"  퍼널: 전체 {len(pdf)} → 공급기관 {int(m_org.sum())} "
        f"→ 제출년도>={year_min} {int(m.sum())}")
    log(f"  특허 1건 이상: 특허DB {int((m & m_nice).sum())} · "
        f"과제메타 {int((m & m_pm).sum())} · 둘 다 {int((m & m_nice & m_pm).sum())}")
    m = m & {"nice": m_nice, "pmeta": m_pm, "both": m_nice & m_pm}[patent_src]
    if allow_subject:                            # 기존 파이프라인과 동일한 연구수행주체 필터
        before = int(m.sum())
        m &= pdf["연구수행주체"].isin(ALLOW).values
        log(f"  연구수행주체 {sorted(ALLOW)} 적용: {before} → {int(m.sum())}")
    if field and field != "전체":                 # 의료·제약 분야로 한정
        import gen_report as _gr
        before = int(m.sum())
        keep = np.zeros(len(m), dtype=bool)
        why = []
        for i in np.flatnonzero(m):
            d = str(pdf["과제설명문"].iat[i])
            ok, w = in_medical_domain(_gr.extract_class(d) or "",
                                      str(pdf["과제명"].iat[i]), d)
            keep[i] = ok
            if not ok:
                why.append(w.split("·")[0].strip())
        m &= keep
        from collections import Counter as _C
        log(f"  의료·제약 분야 한정: {before} → {int(m.sum())}건 "
            f"(제외 {before - int(m.sum())}건)")
        for w, n in _C(why).most_common(6):
            log(f"      제외 사유 {n:>4}건  {w}")
    idx = np.flatnonzero(m)
    log(f"  코퍼스 확정 {len(idx)}건 (특허 기준 '{patent_src}·{patent_status}'"
        + (f" · 분야 '{field}'" if field and field != "전체" else "") + ")")
    if not len(idx):
        raise SystemExit("[중단] 코퍼스가 비었습니다")

    sub = pdf.iloc[idx]
    import gen_report as _gr2
    c = {
        "field": np.array([_gr2.extract_class(str(d)) or ""
                           for d in sub["과제설명문"].values]),
        "pid": pid_all[idx], "pname": sub["과제명"].astype(str).values,
        "pdesc": sub["과제설명문"].astype(str).values, "pkw": sub["키워드_리스트"].values,
        "promise": sub["유망성점수"].values.astype(np.float32),
        "subject": sub["연구수행주체"].astype(str).values,
        "year": y[idx], "org": org_all[idx],
        "pat": pat_nice[idx], "pat_meta": pat_meta[idx],
        "pap": np.array([mt._to_int(pmeta.get(p, {}).get("논문건수")) for p in pid_all[idx]],
                        dtype=np.int32),
        "resolved": {k: v[0] for k, v in resolved.items()},
        "pmeta": pmeta,
    }
    # 임베딩은 통합 필터보다 먼저 만든다 — 통합 시 같은 마스크로 함께 걸러야 한다.
    _M = np.vstack([np.asarray(e, dtype=np.float32) for e in sub["norm_embed"].values])
    c["M"] = _M / np.linalg.norm(_M, axis=1, keepdims=True)
    # ---- 다년차 과제 통합 --------------------------------------------------
    # 같은 과제의 연차 보고를 한 건으로 묶고, 특허는 전 연차 합집합을 쓴다.
    c["members"] = {str(p): [str(p)] for p in c["pid"]}
    c["rep_of"] = {str(p): str(p) for p in c["pid"]}
    if merge_phase:
        sup_of = np.array([resolved_rev.get(o, o) for o in c["org"]])
        rep_of, members = pmg.build_groups(c["pid"], c["pname"], sup_of, c["year"])
        st = pmg.stats(members)
        log(f"· 다년차 과제 통합: {st['rows']}건 → {st['groups']}건 "
            f"(2연차 이상 {st['multi']}개 그룹 · 통합으로 {st['collapsed']}건 감소 · "
            f"최대 {st['max_size']}연차)")
        keep = np.array([str(p) in members for p in c["pid"]])
        # isinstance(v, np.ndarray) 로 걸러선 안 된다 — pid·pname·pdesc·subject 는
        # pandas 문자열 컬럼의 .values 라서 ExtensionArray 이고 ndarray 가 아니다.
        # 길이만 맞으면 불리언 마스크로 자른다(dict 는 제외).
        for k, v in list(c.items()):
            if (not isinstance(v, dict) and hasattr(v, "__len__")
                    and len(v) == len(keep)):
                c[k] = v[keep]
        if len(c["pid"]) != int(keep.sum()):
            raise SystemExit(f"[중단] 연차 통합 필터 적용 실패: pid {len(c['pid'])} "
                             f"≠ {int(keep.sum())}")
        c["members"], c["rep_of"] = members, rep_of
        # 특허·논문: 특허는 전 연차 출원번호 합집합, 논문·과제메타 특허는 연차 최대값
        psets = patent_sets(alive_only=(patent_status == "생존"))
        merged, pm_max, pap_max = [], [], []
        for p in c["pid"]:
            mem = members[str(p)]
            u = set()
            for q in mem:
                u |= psets.get(q, set())
            merged.append(len(u))
            pm_max.append(max((mt._to_int(pmeta.get(q, {}).get("특허건수")) for q in mem),
                              default=0))
            pap_max.append(max((mt._to_int(pmeta.get(q, {}).get("논문건수")) for q in mem),
                               default=0))
        before = int(c["pat"].sum())
        c["pat"] = np.array(merged, dtype=np.int32)
        c["pat_meta"] = np.array(pm_max, dtype=np.int32)
        c["pap"] = np.array(pap_max, dtype=np.int32)
        c["nphase"] = np.array([len(members[str(p)]) for p in c["pid"]], dtype=np.int32)
        c["phase_pids"] = np.array([",".join(members[str(p)]) for p in c["pid"]])
        log(f"  특허 합집합 적용: 대표 행 기준 {before}건 → {int(c['pat'].sum())}건 "
            f"(중위 {int(np.median(c['pat']))}건 · 최대 {int(c['pat'].max())}건)")
        log(f"  통합된 과제 {int((c['nphase'] > 1).sum())}건 · 코퍼스 {len(c['pid'])}건")
    else:
        c["nphase"] = np.ones(len(c["pid"]), dtype=np.int32)
        c["phase_pids"] = np.array([str(p) for p in c["pid"]])

    # 과제명 정상성은 **배제하지 않고 표시만** 한다. 예외 판정(적합도·유망특허)이
    # LLM 채점 이후에만 가능하므로, 후보에는 넣고 최종 선정에서 건너뛴다.
    nb = np.array([bad_project_name(x) for x in c["pname"]], dtype=object)
    c["namebad"] = np.array([bool(x[0]) for x in nb])
    c["namewhy"] = np.array([x[1] for x in nb])
    if check_name and c["namebad"].any():
        from collections import Counter as _C2
        log(f"· 과제명 정상성: {int(c['namebad'].sum())}건 표시"
            f"(연구 주제가 아닌 과제 — 예외 조건 미충족 시 최종 선정에서 제외)")
        for w, n in _C2(c["namewhy"][c["namebad"]]).most_common():
            log(f"      {w} {n}건")
        for nmx, n in _C2(c["pname"][c["namebad"]]).most_common(6):
            log(f"        · {str(nmx)[:44]}{f' ×{n}' if n > 1 else ''}")
    log(f"· 유망 등록특허 집계… (등록 ∧ 유망성 ≥ {promise_min})")
    pms = promising_patent_sets(promise_min)
    c["pat_promise"] = np.array(
        [len(set().union(*[pms.get(q, set()) for q in c["members"].get(str(p), [str(p)])]))
         for p in c["pid"]], dtype=np.int32)
    log(f"  유망 등록특허 1건 이상 과제 {int((c['pat_promise'] >= 1).sum())}"
        f"/{len(c['pid'])}건")
    c["check_name"] = bool(check_name)
    del pdf, sub
    log(f"  코퍼스 준비 완료 ({time.time()-t0:.0f}s) · "
        f"기관 {len(set(c['org']))}곳 · 특허 중위 {int(np.median(c['pat']))}건 · "
        f"유망성 평균 {c['promise'].mean():.1f}")
    return c


# ---------------------------------------------------------------- LLM 재랭킹
# 이 보고서의 기본 방향은 '공급기관(대학·병원·연구기관)이 보유한 과제 기술 → 기업'이다.
# 방향을 명시하지 않으면 모델이 읽기 쉬운 쪽으로 서술해 버려, 기업이 이미 가진 기술을
# 과제에 제공하는 관계까지 '기술도입'처럼 적는다(젠스퀘어·비엠에이 사례). 그래서
# 채점 단계에서 방향(연계유형)을 함께 받아 근거 생성과 보고서 표기에 그대로 쓴다.
_SYS = ("당신은 기업의 기술과 국가 R&D 과제의 연계 가능성을 평가하는 기술이전 전문가입니다. "
        "각 과제가 해당 기업에 얼마나 적합한 이전·협력 대상인지 냉정하게 평가합니다.")

# 연계유형 — 보고서에 그대로 표기한다.
# 기업이 과제에 부품·서비스를 제공하는 '사업화파트너'는 두지 않는다(요청). 이 보고서는
# 공급기관 기술을 기업이 받아 쓰는 기술이전을 다루므로, 방향이 반대인 관계는 유형으로
# 인정하지 않는다. 결과적으로 그런 건도 기술도입/공동연구 중 하나로 판정된다.
KIND_IN, KIND_CO, KIND_NA = "기술도입", "공동연구", "미분류"
KINDS = (KIND_IN, KIND_CO)
KIND_DESC = {
    KIND_IN: "공급기관 과제의 기술을 기업이 이전받아 자사 제품·서비스에 적용",
    KIND_CO: "양측 기술을 결합해 기업과 공급기관이 함께 개발",
    KIND_NA: "유형 판정 없음",
}

# 유형은 '표기'가 목적이므로 순위에는 쓰지 않는다. 유형을 점수에 섞으면 방향을 고르는
# 판정과 적합도 채점이 서로를 흔든다(실측: 유형 항목을 채점 프롬프트에 넣자 같은 건의
# 적합도가 85→40 으로 내려갔다). 그래서 채점은 그대로 두고, 선정된 추천에만 유형을 매긴다.

# 질의 근거별 평가 관점. 근거가 둘이면 관점을 함께 제시하고 점수는 하나만 받는다
# (둘 중 어느 쪽으로든 기여하면 적합한 것으로 본다).
_ASPECT = {
    "수요기술": ("[확보하려는 수요기술]", "기업이 확보하려는 수요기술을 충족·해결"),
    "기보유기술": ("[이미 보유한 기술]", "기업이 이미 보유한 기술을 보강·고도화하거나 "
                          "새 응용으로 확장"),
    "기업DB": ("[기업 사업내용]", "기업의 사업·제품 영역에 이전·접목되어 기술 경쟁력을 높임"),
}


def _guide(u):
    bases = u["근거"]
    labels = " · ".join(_ASPECT[b][0] for b in bases)
    expl = " / ".join(_ASPECT[b][1] for b in bases)
    tail = ("두 관점 중 어느 쪽으로든 기여할 수 있으면 그 정도를 점수에 반영하세요.\n"
            if len(bases) > 1 else "")
    return (f"각 과제를 아래 관점에서 0~100 정수로 평가하세요.\n"
            f"  적합도: 이 과제가 {expl} 할 수 있는 정도 (기준: {labels})\n"
            + tail + mt._BANDS
            + "키워드 표면 일치가 아니라 기술 내용·적용 제품 관점의 실질 적합도를 보세요.\n"
            "반드시 아래 JSON 형식으로만, 모든 과제에 대해 답하세요(reason 은 40자 이내 한국어):\n"
            '{"results": [{"id": <과제번호>, "적합도": <0-100 정수>, "reason": "<근거>"}, ...]}')


def _unit_block(u):
    b = [f"[기업] {u['기업명']}" + (f" ({u['기관유형']})" if u["기관유형"] else "")]
    for base in u["근거"]:
        if base == "수요기술":
            b.append("[확보하려는 수요기술]\n"
                     f"- 기술명: {u['수요기술명']}\n- 내용: {u['수요기술 내용'][:DEM_MAX]}")
            continue
        body = mt._HOLD_NAME_SEC.sub("", u["기보유기술 내용"]).strip()
        head = ("[이미 보유한 기술·사업내용]" if base == "기보유기술" else "[기업 사업내용]")
        meta = " / ".join(x for x in (u["기술유형"], u["기술분야"]) if x)
        b.append(head + "\n"
                 + (f"- 기술명: {u['기보유기술명']}\n" if u["기보유기술명"] else "")
                 + (f"- 구분: {meta}\n" if meta else "")
                 + f"- 내용: {body[:HOLD_MAX]}")
    return "\n\n".join(b)


def _norm_kind(v):
    """모델이 돌려준 유형 문자열 → 정규 유형. 못 알아보면 '미분류'로 드러낸다.

    없는 유형(예: 예전의 '사업화파트너')을 억지로 둘 중 하나에 끼워 넣지 않는다 —
    조용히 잘못 붙는 것보다 검증에서 미분류로 걸리는 편이 낫다.
    """
    t = re.sub(r"[\s·형]", "", str(v or ""))
    for k in KINDS:
        if re.sub(r"[\s·형]", "", k) in t:
            return k
    if "도입" in t or "이전" in t:
        return KIND_IN
    if "공동" in t or "결합" in t:
        return KIND_CO
    return KIND_NA


def _parse(text):
    """{"results":[...]} → {id: (적합도, reason)}."""
    m = re.search(r'\{.*"results".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("results", []) if isinstance(data, dict) else data
    except Exception:
        items = [json.loads(x) for x in re.findall(r'\{[^{}]*"id"[^{}]*\}', blob)
                 if mt._loadable(x)]
    out = {}
    for it in items:
        try:
            i = int(it["id"])
        except (KeyError, ValueError, TypeError):
            continue
        out[i] = (mt._clip(it.get("적합도")), str(it.get("reason", "")).strip()[:120])
    return out


def score_candidates(u, cands, retry=mt.RETRY):
    """후보를 LLM_BATCH 단위로 채점 → {id: (적합도, reason)}. 누락 id 는 재질의."""
    guide = _guide(u)
    scores, miss = {}, 0
    for b in range(0, len(cands), cm.LLM_BATCH):
        batch = cands[b:b + cm.LLM_BATCH]
        parsed = _score_batch(u, batch, guide)
        for _ in range(retry):
            left = [x for x in batch if x["id"] not in parsed]
            if not left:
                break
            parsed.update(_score_batch(u, left, guide))
        for x in batch:
            if x["id"] in parsed:
                scores[x["id"]] = parsed[x["id"]]
            else:
                scores[x["id"]] = (0, "(LLM 응답 누락)")
                miss += 1
    if miss:
        log(f"      ! LLM 응답 누락 {miss}/{len(cands)}건(재질의 {retry}회 후에도 없음 → 0점)")
    return scores


def _score_batch(u, batch, guide):
    user = f"{_unit_block(u)}\n\n{mt._cand_block(batch)}\n\n{guide}"
    msgs = [{"role": "system", "content": _SYS}, {"role": "user", "content": user}]
    try:
        return _parse(cm.stream_explanation(msgs, max_tokens=cm.LLM_MAXTOK,
                                            temperature=0.0, top_p=1.0))
    except Exception as e:
        log(f"      ! LLM 배치 실패: {e}")
        return {}


# ---------------------------------------------------------------- 연계유형 판정
_KIND_SYS = ("당신은 대학·병원·연구기관의 국가 R&D 과제와 기업을 잇는 기술이전 전문가입니다. "
             "과제와 기업 사이 협력의 '방향'을 근거를 보고 냉정하게 판정합니다.")


def _kind_guide(n):
    return (f"아래 과제 {n}건에 대해, 이 기업과의 협력 방향을 판정하세요.\n"
            "이 보고서는 공급기관(대학·병원·연구기관)이 보유한 과제 기술을 기업이 받아 쓰는 "
            "기술이전을 다룹니다. 둘 중 하나를 고르세요.\n"
            f"  ① {KIND_IN}: 기업이 과제의 기술을 넘겨받아 자사 제품·서비스에 바로 적용할 수 "
            "있는 관계.\n"
            f"  ② {KIND_CO}: 어느 한쪽 기술만으로는 부족해 양측 기술을 결합해 새로 개발해야 "
            "하는 관계.\n"
            "판정 기준\n"
            "  · 과제의 기술 요소가 기업이 필요로 하는 것을 그대로 채워 주면 "
            f"→ {KIND_IN}\n"
            "  · 과제 기술만으로는 기업 제품이 되지 않고 기업 기술과 맞물려야 결과가 나오면 "
            f"→ {KIND_CO}\n"
            "  · 기업이 이미 비슷한 역량을 갖고 있어도, 과제에서 얻을 기술 요소가 그 역량을 "
            f"보강·고도화한다면 {KIND_IN} 으로 본다.\n"
            "  · 기업의 사업이 중개·유통·용역이어서 과제 기술을 자사 서비스의 대상·자산으로 "
            f"다루는 경우도, 그 기술을 확보해 쓰는 것이므로 {KIND_IN} 으로 본다.\n"
            "반드시 아래 JSON 형식으로만, 모든 과제에 답하세요(이유는 30자 이내):\n"
            '{"results": [{"id": <번호>, "유형": "'
            + f'{KIND_IN}|{KIND_CO}", "이유": "<근거>"}}, ...]}}')


def _parse_kind(text):
    m = re.search(r'\{.*"results".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("results", []) if isinstance(data, dict) else data
    except Exception:
        items = [json.loads(x) for x in re.findall(r'\{[^{}]*"id"[^{}]*\}', blob)
                 if mt._loadable(x)]
    out = {}
    for it in items:
        try:
            i = int(it["id"])
        except (KeyError, ValueError, TypeError):
            continue
        out[i] = (_norm_kind(it.get("유형")), str(it.get("이유", "")).strip()[:60])
    return out


def classify_kinds(u, sel, c, ckpt=None, retry=1):
    """선정된 추천들의 연계유형을 한 번에 판정 → {pid: (유형, 이유)}.

    채점(적합도)과 분리한 이유: 유형 항목을 채점 프롬프트에 함께 넣자 같은 건의 적합도가
    85→40 으로 흔들렸다. 채점은 후보 120건에 걸고, 유형은 최종 선정분에만 매긴다
    (기업당 1회 호출이라 값도 싸다).
    """
    out, need = {}, []
    for x in sel:
        pid = str(c["pid"][x["i"]])
        ck = f"{u['키']}::{pid}"
        if ckpt is not None and ck in ckpt:
            out[pid] = tuple(ckpt[ck])
        else:
            need.append((pid, x))
    if not need:
        return out
    lines = ["[평가 대상 R&D 과제]"]
    for j, (pid, x) in enumerate(need, 1):
        i = x["i"]
        lines.append(f"{j}. 과제명: {c['pname'][i]}\n   설명: {c['pdesc'][i][:cm.DESC_CAND]}")
    msgs = [{"role": "system", "content": _KIND_SYS},
            {"role": "user", "content": f"{_unit_block(u)}\n\n" + "\n".join(lines)
                                        + f"\n\n{_kind_guide(len(need))}"}]
    got = {}
    for _ in range(retry + 1):
        try:
            got = _parse_kind(cm.stream_explanation(msgs, max_tokens=700, temperature=0.0,
                                                   top_p=1.0))
        except Exception as e:
            log(f"      ! 연계유형 판정 실패: {e}")
        if len(got) >= len(need):
            break
    for j, (pid, _x) in enumerate(need, 1):
        kind, why = got.get(j, (KIND_NA, ""))
        out[pid] = (kind, why)
        if ckpt is not None:
            ckpt[f"{u['키']}::{pid}"] = [kind, why]
    return out


# ---------------------------------------------------------------- 유형별 근거문
# xp.gen_reason 은 방향을 정하지 않아 모델이 읽기 쉬운 쪽으로 쓴다. 여기서는 채점 단계가
# 정한 연계유형에 맞춰 '누가 무엇을 받는지'를 프롬프트로 고정한다. 금지표현·다듬기·
# 유망성 차단은 explain_meditek_top10 의 것을 그대로 재사용한다.
_DIR_RULE = {
    KIND_IN: ("주어는 '과제의 기술'이다. 과제가 보유한 기술을 이 기업이 이전받아 자사 "
              "제품·서비스에 적용하는 방향으로만 써라. 기업의 기술이 과제에 쓰인다는 "
              "서술은 절대 하지 마라."),
    KIND_CO: ("양측 기술을 결합하는 방향으로 써라. 어느 한쪽이 일방적으로 제공하는 것이 "
              "아니라 함께 개발해야 하는 관계임이 드러나야 한다."),
}
_DIR_SHOT = {
    KIND_IN: [("기술명: 디지털 후각·뇌혈류 기반 인지건강 선별 솔루션",
               "과제명: 광학기술기반 뇌 혈류 산소포화도 측정 연구\n"
               "과제설명: 근적외선 분광법으로 전두엽 혈류를 계측하는 진단장비 개발",
               "근적외선 분광 기반 뇌혈류 측정 기술을 도입해 당사 인지건강 선별 솔루션의 "
               "계측 정확도를 높일 수 있음")],
    KIND_CO: [("기술명: 오가노이드 3D 배양·이미징 통합 자동화 플랫폼",
               "과제명: 3차원 혈관 뇌 장벽 생체 조직 칩 제품 기술 개발\n"
               "과제설명: 고품질 BBB 조직칩 제조와 약동학 분석",
               "조직칩 제조 기술과 당사 자동화 분주·이미징 플랫폼을 결합해 약물 평가 "
               "일괄 시스템으로 공동개발 가능")],
}


def _mr_sys(kind):
    return (
        "너는 대학·연구기관의 국가 R&D 과제와 기업을 잇는 '매칭 근거'를 한 문장으로 쓰는 "
        "한국어 AI다.\n"
        f"이 건의 연계유형은 '{kind}' 이다: {KIND_DESC.get(kind, '')}\n"
        f"**방향 규칙: {_DIR_RULE.get(kind, '')}**\n"
        "작성 규칙:\n"
        "1) 70자 이내 한 문장. 명사형/음슴체 종결.\n"
        "2) 과제 개요·수행기관·기간 같은 '설명'은 쓰지 말고 두 대상의 접점만 쓴다.\n"
        "3) 부정 평가('~ 부재', '~ 미흡', '불일치')는 쓰지 않는다. 완전히 일치하지 않아도 "
        "기여 가능성으로 표현한다.\n"
        "4) 문장 하나만 출력. 따옴표·머리기호·부연 금지.\n"
        "5) few-shot 은 방향과 톤 참고용이며 고유명사를 재사용하지 않는다.\n"
        "6) [과제 보유 특허] 가 주어지고 그 특허가 기업 기술과 실제로 맞물릴 때는 특허명을 "
        "근거로 인용해 접점을 구체화한다. 맞물리지 않으면 언급하지 않는다 — 억지로 끼워 "
        "넣지 않는다. 인용할 때는 **주어진 표기를 한 글자도 바꾸지 말고 그대로 쓰고 "
        "작은따옴표로 감싼다**('…'). 줄이거나 풀어 쓰지 않는다. 목록에 없는 특허명을 "
        "만들어 쓰지 않는다.\n"
        "7) " + xp._BAN_RULE)


def gen_reason_typed(u, proj, kind, retry=2):
    """연계유형에 맞는 방향으로 표용 한 문장 근거를 만든다."""
    blocks = []
    if u["수요기술 내용"]:
        blocks.append(f"기술명: {u['수요기술명']}\n기술 내용: {u['수요기술 내용'][:600]}")
    body = xp.hold_body(u)
    if body:
        blocks.append(f"기술명: {u['기보유기술명']}\n기술 내용(보유): {body[:600]}")
    pj = f"과제명: {proj.get('과제명','')}\n과제설명: {str(proj.get('설명','') or '')[:600]}"
    pats = (proj.get("특허명") or [])[:3]
    if pats:
        pj += "\n과제 보유 특허(질의와 유사한 것만): " + " / ".join(pats)
    shots = []
    for tech, prj, ans in _DIR_SHOT.get(kind, []):
        shots += [{"role": "user",
                   "content": f"[기업 기술]\n{tech}\n\n[R&D 과제]\n{prj}\n\n매칭 근거(한 문장):"},
                  {"role": "assistant", "content": ans}]
    user = ("[기업 기술]\n" + "\n\n".join(blocks)
            + f"\n\n[R&D 과제]\n{pj}\n\n매칭 근거(한 문장):")
    msgs = [{"role": "system", "content": _mr_sys(kind)}] + shots + \
           [{"role": "user", "content": user}]
    out = ""
    for _ in range(retry + 1):
        try:
            raw = cm.stream_explanation(msgs, max_tokens=140, temperature=0.0,
                                        top_p=1.0).strip()
        except Exception as e:
            log(f"      ! 매칭근거 실패: {e}")
            return ""
        out = raw.strip().strip('"').strip("'").split("\n")[0].strip()
        if out.startswith("매칭 근거"):
            out = out.split(":", 1)[-1].strip()
        out = cm.normalize_spacing(out[:150])
        if not xp.BANNED.search(out):
            return out
    return out


# 상세근거의 섹션 가이드는 xp 에서 '근거 종류'(수요/보유)로만 갈린다. 둘 다 '과제 → 기업'
# 방향을 전제하므로, 공동연구 건에서는 방향 규칙과 어긋날 수 있다
# (가이드는 "과제가 수요를 충족·해결하는 방식"을 쓰라 하고 규칙은 그 반대를 금지한다).
# 그래서 유형별로 '연관성'·'기술 적합성' 두 섹션의 지시를 바꿔 끼운다.
# '추천 과제의 우수성'은 과제 자체를 서술하므로 그대로 둔다.
_SEC_BY_KIND = {
    KIND_CO: {
        "연관성": ("company 의 기술과 과제의 기술이 각각 무엇을 담당하는지 나눠 짚고, 둘이 "
                "만나는 지점을 설명. 어느 한쪽이 일방적으로 제공하는 관계로 쓰지 말 것"),
        "기술 적합성": ("과제의 기술과 이 기업의 기술을 어느 지점에서 결합해야 하는지, 결합해야만 "
                   "얻어지는 결과가 무엇인지 설명. 각자 부족한 부분을 상대가 채우는 구조를 "
                   "구체적으로 쓰고, 한쪽이 다른 쪽에 기술을 넘겨준다는 서술은 하지 말 것"),
    },
}


def gen_detail_typed(u, kws, proj, kind, retry=2):
    """3섹션 상세근거 — 섹션 지시와 방향 규칙을 모두 유형에 맞춘다."""
    p, fmt = xp.payload_for(u, kws, proj)
    guide = dict(p["output_requirements"]["section_guide"])
    guide.update(_SEC_BY_KIND.get(kind, {}))
    p["output_requirements"]["section_guide"] = {k: guide[k] for k in fmt}
    p["output_requirements"]["연계유형"] = kind
    p["output_requirements"]["연계유형_설명"] = KIND_DESC.get(kind, "")
    msgs = cm.build_messages(p, direction="company")
    msgs[-1]["content"] += ("\n" + xp._BAN_RULE
                            + f"\n**이 건의 연계유형은 '{kind}' 이다: "
                              f"{KIND_DESC.get(kind, '')}\n{_DIR_RULE.get(kind, '')}**")
    if p.get("related_research", {}).get("patent"):
        # related_research.patent 에는 **질의와 유사한 특허만** 들어 있다. 실제로 기업
        # 기술과 맞물릴 때만 특허명을 인용하고, 아니면 언급하지 않게 지시한다.
        msgs[-1]["content"] += (
            "\n**related_research.patent / related_research.paper 는 이 과제가 확보한 "
            "특허·논문 중 이 기업의 기술과 의미가 가까운 것만 골라 넣은 것이다. 실제로 "
            "맞물리는 지점이 있으면 '연관성'과 '유사 사례 및 실적' 에서 그 이름을 인용해 "
            "무엇을 이전받아 어디에 쓰는지 적어라. 맞물리지 않으면 언급하지 마라.\n"
            "인용 규칙(반드시 지켜라):\n"
            "  · 이름은 **목록의 표기를 한 글자도 바꾸지 말고 그대로** 쓴다. 줄이거나 "
            "풀어 쓰거나 요약하지 않는다.\n"
            "  · 이름은 **작은따옴표로 감싼다**('…'). 감싸지 않은 채 문장에 녹여 쓰지 "
            "않는다 — 예를 들어 \"특허 1건이 존재하여 A와 같은 기반이 있다\" 처럼 쓰지 "
            "말고 \"'A' 특허를 확보했다\" 처럼 쓴다.\n"
            "  · 목록에 없는 이름을 만들어 쓰지 않는다.**")
    text = ""
    for i in range(retry + 1):
        out = cm.stream_explanation(msgs, max_tokens=1400,
                                    temperature=0.2 if i == 0 else 0.0,
                                    top_p=0.9, expected_keys=fmt)
        secs = cm.parse_sections(out, tuple(fmt))
        parts = [f"[{k}] {secs.get(k, '').strip()}" for k in fmt if secs.get(k, "").strip()]
        text = cm.normalize_spacing("\n\n".join(parts) if parts else out.strip())
        if not xp.BANNED.search(text):
            return text
    return text

# ---------------------------------------------------------------- 행 생성
_PAPERS_JSON = None


def _papers_of(pid):
    """보고서 논문 표와 같은 목록(pid_papers.json). 없으면 빈 목록."""
    global _PAPERS_JSON
    if _PAPERS_JSON is None:
        fp = os.path.join(os.environ.get("COMPA_SCRATCH", cm.OUT_DIR),
                          "pid_papers.json")
        try:
            _PAPERS_JSON = json.load(open(fp, encoding="utf-8"))
        except Exception:
            log(f"  (참고) {fp} 없음 — 근거문 논문 건수는 과제메타 값을 쓴다")
            _PAPERS_JSON = {}
    return _PAPERS_JSON.get(str(pid), [])


def make_row(u, x, rank, c, rev, ex_ck, rs_ck, ex_path, rs_path, dry_run, no_detail,
             kind=None, kind_why=""):
    """선정된 후보 하나 → 산출 행. 근거문이 없으면 여기서 35B 로 만들어 캐시에 넣는다."""
    i = x["i"]
    pid = str(c["pid"][i])
    proj = {
        "pid": pid, "과제명": c["pname"][i], "설명": c["pdesc"][i],
        "수행기관": c["org"][i], "키워드": [],
        # 논문은 **보고서 표와 같은 소스**(pid_papers.json)를 쓴다. NTIS 과제메타의
        # 논문건수는 연차 합산·중복 처리 규칙이 달라 표와 어긋난다 — 실측에서 근거문의
        # 논문 건수 오류 39건이 전부 이 불일치 때문이었다(특허는 0건).
        "논문명": [x.get("논문명", "") for x in _papers_of(pid)],
        # 질의와 실제로 유사한 특허만 넣는다. 관련 특허가 없으면 빈 목록이 되고,
        # 프롬프트 규칙상 그때는 특허를 언급하지 않는다(억지 인용 방지).
        "특허명": [t for t, _st, _s in (x.get("pat_titles") or [])],
        "논문건수": len(_papers_of(pid)) or int(c["pap"][i]),
        "특허건수": int(c["pat"][i]),
        "총연구비_상위비율": c["pmeta"].get(pid, {}).get("총연구비_상위비율"),
        "논문건수_상위비율": c["pmeta"].get(pid, {}).get("논문건수_상위비율"),
    }
    xp.assert_no_promise(proj, "근거 프롬프트")
    kind = kind or x.get("kind") or KIND_NA
    ckey = f"{u['번호'] or u['키']}::{pid}"
    if not dry_run:
        if ckey not in rs_ck:
            rs_ck[ckey] = gen_reason_typed(u, proj, kind)
            json.dump(rs_ck, open(rs_path, "w", encoding="utf-8"), ensure_ascii=False)
        if not no_detail and ckey not in ex_ck:
            ex_ck[ckey] = gen_detail_typed(u, u["키워드"], proj, kind)[:cm.DESC_OUT]
            json.dump(ex_ck, open(ex_path, "w", encoding="utf-8"), ensure_ascii=False)
    return {
        "번호": u["번호"], "관리번호": u["관리번호"], "기업명": u["기업명"],
        "사업자등록번호": u["사업자등록번호"], "기관유형": u["기관유형"],
        "소스": u["소스"], "질의길이": u["질의길이"],
        "기보유기술_보유": u["기보유기술_보유"], "기업DB_기업명": u["기업DB_기업명"],
        "수요기술명": u["수요기술명"], "기보유기술명": u["기보유기술명"],
        "키워드": cm.SEP.join(u["키워드"]),
        "rank": rank, "순위구분": x["순위구분"], "연계유형": kind,
        "연계유형근거": kind_why,
        "최종점수": round(x["total"], 2), "적합도": x["fit"],
        "1차점수": round(100 * x["pre"], 2),
        "유사도_코사인": round(x["cos"], 6), "유사도_정규화": round(x["cos_n"], 4),
        "유망성점수": round(float(c["promise"][i]), 2),
        "특허축점수": round(x.get("pat", 0.0), 4),
        "유사특허수": int(x.get("pat_cnt", 0)),
        "특허적합에너지": round(x.get("pat_e", 0.0), 4),
        "특허유사강도": round(x.get("pat_sim_e", 0.0), 4),
        "특허권리상태": round(x.get("pat_legal", 0.0), 4),
        "특허최신성": round(x.get("pat_recent", 0.0), 4),
        "특허최고유사도": round(x.get("pat_max", 0.0), 4),
        "과제권리최고": round(x.get("pat_legal_max", 0.0), 4),
        "유망등록특허수": int(c["pat_promise"][i]),
        "인용특허명": " | ".join(f"{t}({st})" for t, st, _ in (x.get("pat_titles") or [])),
        "과제명특이": bool(c["namebad"][i]),
        "과제명특이사유": str(c["namewhy"][i]),
        "특허임계값": round(x.get("pat_tau", 0.0), 4),
        "특허건수": int(c["pat"][i]), "특허건수_과제메타": int(c["pat_meta"][i]),
        "연차수": int(c["nphase"][i]), "연차_과제고유번호": str(c["phase_pids"][i]),
        # 논문건수는 **보고서 논문 표와 같은 소스**를 쓴다(특허건수와 같은 방침).
        # NTIS 과제메타 값은 연차 합산 규칙이 달라 표와 어긋나므로 별도 컬럼에 남긴다.
        "논문건수": len(_papers_of(pid)) or int(c["pap"][i]),
        "논문건수_과제메타": int(c["pap"][i]),
        "과제고유번호": pid, "과제명": c["pname"][i],
        "과제분야": c["field"][i],
        "과제수행기관": c["org"][i], "공급기관": rev.get(c["org"][i], ""),
        "제출년도": int(c["year"][i]) if np.isfinite(c["year"][i]) else 0,
        "연구수행주체": c["subject"][i],
        "판단근거": xp.polish(rs_ck.get(ckey) or x["reason"]),
        "추천근거_상세": xp.polish(ex_ck.get(ckey, "")),
        "LLM채점근거": x["reason"],
        "과제설명문": c["pdesc"][i][:cm.DESC_OUT],
    }


# ---------------------------------------------------------------- 산출
OUT_COLS = ["번호", "관리번호", "기업명", "사업자등록번호", "기관유형", "소스", "질의길이",
            "기보유기술_보유", "기업DB_기업명", "수요기술명", "기보유기술명", "키워드",
            "rank", "순위구분", "연계유형", "연계유형근거", "최종점수", "적합도", "1차점수", "유사도_코사인", "유사도_정규화",
            "유망성점수", "특허축점수", "유사특허수", "특허적합에너지", "특허유사강도",
            "특허권리상태", "특허최신성", "특허최고유사도", "특허임계값", "과제권리최고",
            "유망등록특허수", "인용특허명", "과제명특이", "과제명특이사유",
            "특허건수", "특허건수_과제메타", "논문건수", "논문건수_과제메타",
            "연차수", "연차_과제고유번호",
            "과제고유번호", "과제명", "과제분야", "과제수행기관", "공급기관", "제출년도", "연구수행주체",
            "판단근거", "추천근거_상세", "LLM채점근거", "과제설명문"]
XLSX_WIDTH = dict(mt.XLSX_WIDTH, **{
    "사업자등록번호": 14, "질의길이": 8, "기보유기술_보유": 11, "기업DB_기업명": 16,
    "1차점수": 8, "유사도_정규화": 11, "특허건수_과제메타": 13, "공급기관": 22,
    "논문건수_과제메타": 13, "특허축점수": 10, "유사특허수": 9, "특허적합에너지": 13, "특허유사강도": 11,
    "특허권리상태": 11, "특허최신성": 10, "특허최고유사도": 13, "특허임계값": 10,
    "과제권리최고": 12, "유망등록특허수": 13, "과제명특이": 10, "과제명특이사유": 18,
    "인용특허명": 46,
    "연차수": 7, "연차_과제고유번호": 30,
    "과제분야": 16, "연계유형": 11, "연계유형근거": 28,
    "제출년도": 8, "연구수행주체": 11, "판단근거": 40, "추천근거_상세": 60, "LLM채점근거": 30})
WRAP = mt.WRAP_COLS | {"판단근거", "추천근거_상세", "LLM채점근거", "공급기관"}


def save_xlsx(df, path):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    import openpyxl
    df.to_excel(path, index=False)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    for j, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(j)].width = XLSX_WIDTH.get(col, 14)
        if col in WRAP:
            for i in range(2, ws.max_row + 1):
                ws.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    wb.save(path)


def org2supply(resolved):
    """과제수행기관명 → 공급기관명(역인덱스). 겹치면 ' / ' 로 잇는다."""
    rev = {}
    for sup, orgs in resolved.items():
        for o in orgs:
            rev[o] = f"{rev[o]} / {sup}" if o in rev else sup
    return rev


def report_json(df, units, out):
    """gen_report_meditek_top10 과 같은 모양의 보고서 입력 JSON."""
    ub = {u["기업명"]: u for u in units}
    best = {}
    for name, g in df.groupby("기업명", sort=False):
        u = ub[name]
        key = u["번호"] or u["관리번호"] or cm.norm_name(name)
        tops = []
        for r in g.sort_values("rank").to_dict("records"):
            tops.append({
                "rank": int(r["rank"]), "과제고유번호": str(r["과제고유번호"]),
                "과제명": r["과제명"], "수행기관": r["과제수행기관"],
                "공급기관": r["공급기관"], "연계유형": r.get("연계유형", ""),
                "연계유형근거": r.get("연계유형근거", ""),
                "적합도": int(r["적합도"]),
                "최종점수": float(r["최종점수"]), "특허건수": int(r["특허건수"]),
                "논문건수": int(r["논문건수"]), "순위구분": r["순위구분"],
                "판단근거": r["판단근거"], "추천근거_상세": r["추천근거_상세"],
                "과제설명문": r["과제설명문"],
            })
        best[key] = {
            "번호": u["번호"], "관리번호": u["관리번호"], "기업명": name,
            "사업자등록번호": u["사업자등록번호"], "소스": u["소스"],
            "키워드": cm.SEP.join(u.get("키워드", [])),
            "수요기술명": xp.polish(u["수요기술명"]), "수요기술 내용": xp.polish(u["수요기술 내용"]),
            "기보유기술명": xp.polish(u["기보유기술명"]),
            "기보유기술 내용": xp.polish(xp.hold_body(u)),
            "기술유형": u["기술유형"], "기술분야": u["기술분야"], "top5": tops,
        }
    xp.assert_no_promise(best, "보고서 JSON")
    json.dump(best, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return best


# ---------------------------------------------------------------- 검증
def verify(df, units, skipped, c, final, patent_src, field="의료제약",
           patent_status="생존"):
    """산출물 자기검증 — 요구사항별로 통과/실패를 표로 남긴다."""
    ok_org = set().union(*c["resolved"].values()) if c else set()
    pat_ok = {str(p): int(n) for p, n in zip(c["pid"], c["pat"])} if c else {}
    res, warn = [], []

    def chk(name, cond, detail=""):
        res.append((bool(cond), name, detail))

    chk("추천 행 존재", len(df) > 0, f"{len(df)}행")
    n_co = df["기업명"].nunique()
    chk("대상 기업 = 질의 구성 성공 기업", n_co == len(units), f"{n_co}/{len(units)}개사")
    bad_org = sorted(set(df.loc[~df["과제수행기관"].isin(ok_org), "과제수행기관"]))
    chk("모든 과제가 공급기관 수행", not bad_org, f"위반 기관 {bad_org[:5]}" if bad_org
        else f"기관 {df['과제수행기관'].nunique()}곳 전부 공급기관 소속")
    zero = df[df["특허건수"] < 1]
    chk(f"모든 과제 특허 1건 이상({patent_src}·{patent_status})", zero.empty,
        f"위반 {len(zero)}건" if not zero.empty
        else f"최소 {int(df['특허건수'].min())}건 · 중위 {int(df['특허건수'].median())}건")
    # ---- 특허축 -------------------------------------------------------------
    if "특허축점수" in df.columns and float(df["특허축점수"].abs().sum()) > 0:
        # 0점은 정상값 — 질의 유사도 P90 을 넘는 특허가 그 과제에 없다는 뜻이다.
        # 축이 그런 과제에 대해 침묵하고 다른 축이 순위를 정하도록 설계했다.
        chk("특허축 값 유한", np.isfinite(df["특허축점수"]).all(),
            f"평균 {df['특허축점수'].mean():.3f} · 범위 "
            f"{df['특허축점수'].min():.3f}~{df['특허축점수'].max():.3f} · "
            f"0점 {int(df['특허축점수'].le(1e-9).sum())}/{len(df)}행(유사특허 없음)")
        pv = df[df["유사특허수"] > 0]
        chk("유사특허 보유행은 특허축>0", pv.empty or pv["특허축점수"].gt(0).all(),
            f"위반 {int(pv['특허축점수'].le(0).sum())}행" if not pv.empty
            else "유사특허 보유행 없음")
        z = df[df["유사특허수"] == 0]
        chk("유사특허 없는 행은 최고유사도가 임계 이하",
            z.empty or (z["특허최고유사도"] <= z["특허임계값"] + 1e-6).all(),
            f"{len(z)}행 검사")
        chk("특허축 0~1 범위", df["특허축점수"].between(0, 1).all(),
            f"이탈 {int((~df['특허축점수'].between(0, 1)).sum())}행")
        for col in ("특허권리상태", "특허최신성", "특허유사강도"):
            chk(f"{col} 0~1 범위", df[col].between(0, 1).all(),
                f"평균 {df[col].mean():.3f}")
        # 품질계수 범위 — 곱하기 방식이므로 [0.55, 1.00] 을 벗어날 수 없다
        qmin = pax.W_PSIM
        qmax = pax.W_PSIM + pax.W_PLEGAL + pax.W_PRECENT
        qual = (pax.W_PSIM + pax.W_PLEGAL * df["특허권리상태"]
                + pax.W_PRECENT * df["특허최신성"])
        chk(f"품질계수 {qmin}~{qmax} 범위", qual.between(qmin - 1e-9, qmax + 1e-9).all(),
            f"평균 {qual.mean():.3f} · 범위 {qual.min():.3f}~{qual.max():.3f}")
        chk("유사특허수 ≤ 특허건수", (df["유사특허수"] <= df["특허건수"]).all(),
            f"위반 {int((df['유사특허수'] > df['특허건수']).sum())}행 · "
            f"유사특허 보유 {int(df['유사특허수'].gt(0).sum())}/{len(df)}행")
        # 방향성: 유사특허가 많을수록 특허축이 높아야 한다(기업 내 순위상관)
        import scipy.stats as _st
        rs = [_st.spearmanr(g["유사특허수"], g["특허축점수"]).statistic
              for _, g in df.groupby("기업명") if g["유사특허수"].nunique() > 1]
        rho = float(np.nanmean(rs)) if rs else float("nan")
        chk("유사특허수 ↑ → 특허축 ↑ (기업내 순위상관>0)", not rs or rho > 0,
            f"평균 ρ={rho:+.3f} ({len(rs)}개사)")
        # 권리 생존 여부는 **과제 단위**로 본다. 질의와 가장 가까운 특허가 거절돼도
        # 그 과제에 등록·공개 특허가 있으면 이전 협의는 가능하다. 코퍼스 필터를
        # '출원 1건 이상'으로 유지한 정책상 거절뿐인 과제도 들어올 수 있으므로
        # 실패가 아니라 주의로 남긴다.
        if "과제권리최고" in df.columns:
            dead = df[df["과제권리최고"] < pax.LIVE_MIN]
            chk("추천 과제 권리 생존 판정 산출", df["과제권리최고"].gt(0).all(),
                f"생존 {int(df['과제권리최고'].ge(pax.LIVE_MIN).sum())}/{len(df)}행 · "
                f"등록보유 {int(df['과제권리최고'].ge(1.0).sum())}행")
            chk("모든 추천 과제가 생존 특허 보유(출원 계류 또는 등록)", dead.empty,
                f"위반 {len(dead)}건: " + ", ".join(
                    f"{r['기업명']}({r['rank']}위)" for r in dead.head(3).to_dict("records"))
                if not dead.empty else
                f"등록 {int(df['과제권리최고'].ge(1.0).sum())}행 · "
                f"출원 계류 {int(df['과제권리최고'].between(0.4, 0.99).sum())}행")
    # 과제명이 연구 주제가 아닌 추천은 예외 조건(적합도·유망 등록특허)을 반드시 만족해야 한다
    if "과제명특이" in df.columns:
        oo = [r["과제명"] for r in df.to_dict("records") if org_only_name(r["과제명"])]
        chk("기관명 단독 과제명 추천 없음", not oo,
            f"위반 {len(oo)}건: {sorted(set(oo))[:3]}" if oo else "0건")
        nb = df[df["과제명특이"].astype(bool)]
        viol = nb[(nb["적합도"] < NAME_EXEMPT_FIT) | (nb["유망등록특허수"] < 1)
                  | (nb["과제명특이사유"] == "기관명 단독")]
        chk("과제명 비정상 추천은 예외 조건 충족", viol.empty,
            f"위반 {len(viol)}건: "
            + ", ".join(f"{r['기업명']}({r['과제명'][:14]} 적합도 {r['적합도']}·유망특허 "
                        f"{r['유망등록특허수']})" for r in viol.head(3).to_dict("records"))
            if not viol.empty else
            (f"비정상 과제명 {len(nb)}행 전부 예외 충족(적합도 ≥{NAME_EXEMPT_FIT} ∧ "
             f"유망 등록특허 ≥1): "
             + ", ".join(f"{r['과제명'][:16]}(적합도 {r['적합도']})"
                         for r in nb.drop_duplicates('과제고유번호').to_dict('records')[:3])
             if len(nb) else "비정상 과제명 추천 0행"))
    # ---- 특허·논문 인용 -----------------------------------------------------
    if "인용특허명" in df.columns:
        import meditek_cite_mark as _ck

        def _n(t):
            return _ck._norm(t)          # 공백·따옴표 제거 + 전각/반각 통일

        # 그 과제가 실제로 가진 성과 제목(특허·논문)을 정답으로 쓴다.
        # 프롬프트에 넣은 후보(인용특허명)만 대조하면 논문 인용이 환각으로 잡힌다.
        # 보고서 입력 JSON 은 스크래치 디렉터리에 있다(보고서 생성기와 같은 경로).
        _sc = os.environ.get("COMPA_SCRATCH", cm.OUT_DIR)
        def _load(name):
            fp = os.path.join(_sc, name)
            return json.load(open(fp, encoding="utf-8")) if os.path.exists(fp) else {}
        _pats, _paps = _load("pid_patents.json"), _load("pid_papers.json")
        if not _pats:
            log(f"  (참고) {_sc}/pid_patents.json 없음 — 인용 환각 검사는 "
                "프롬프트 후보 목록만 대조한다")

        # 모델이 인용부호로 감싼 '제목처럼 보이는' 구간만 뽑는다.
        # 따옴표를 순서대로 짝짓는 방식은 못 쓴다 — 영문 아포스트로피나 홀따옴표가
        # 섞이면 짝이 밀려 "'A' 논문과 'B'" 의 가운데('  논문과  ')를 잡는다.
        _JOSA = "의는은이가을를과와로에서도만부터까지"
        _TITLEISH = re.compile(r"장치|방법|시스템|플랫폼|소재|조성물|기술개발|[A-Za-z]{6}")

        def _quoted_spans(t):
            """인용 '주장'으로 볼 구간만. 기업 자기 제품·기술명도 따옴표로 감싸므로
            (예: 'Endentics Rinse', 'AI 기반 다중 전극 … 시스템'), 뒤 30자 안에
            '특허'/'논문'이 붙은 것만 성과 인용으로 본다 — 마커의 판정 규칙과 같다."""
            out = []
            for m in re.finditer(r"['‘\"“『「]\s*([^'’\"”』」\n]{10,110}?)\s*['’\"”』」]",
                                 t):
                if not re.search(r"특허|논문", t[m.end():m.end() + 30]):
                    continue
                v = m.group(1).strip()
                if v[:1] in _JOSA or v[:1].isspace():      # 조사로 시작 = 문장 조각
                    continue
                if re.search(r"다\.|\. |특허는|논문은|특허를|논문을", v):
                    continue                               # 문장이 통째로 들어옴
                if not _TITLEISH.search(v):
                    continue
                out.append(v)
            return out

        cited = df[df["인용특허명"].astype(str).str.len() > 0]
        used, halluc = 0, []
        for r in cited.to_dict("records"):
            names = [x.rsplit("(", 1)[0] for x in str(r["인용특허명"]).split(" | ")]
            body = _n(r.get("추천근거_상세", "")) + _n(r.get("판단근거", ""))
            if any(_n(t)[:12] and _n(t)[:12] in body for t in names):
                used += 1
            pid = str(r["과제고유번호"])
            real = [x["특허명"] for x in _pats.get(pid, [])]
            real += [x.get("논문명", "") for x in _paps.get(pid, [])]
            real = [x for x in real if x] or names
            det = str(r.get("추천근거_상세", ""))
            for t in _quoted_spans(det):
                if len(t) < 10 or not re.search(r"장치|방법|시스템|플랫폼|소재|조성물"
                                                r"|[A-Za-z]{6}", t):
                    continue
                nt = _n(t)
                if not any(nt in _n(x) or _n(x) in nt for x in real):
                    halluc.append((r["기업명"], t[:26]))
        chk("근거문의 인용이 그 과제 성과 목록 안에 있음(환각 없음)", not halluc,
            f"인용행 {len(cited)}행 · 위반 {len(halluc)}건"
            + (f" {halluc[:2]}" if halluc else ""))
        chk("인용후보가 있는 행은 근거문에 실제로 반영",
            not len(cited) or used / len(cited) >= 0.6,
            f"{used}/{len(cited)}행 반영 ({used / max(len(cited), 1) * 100:.0f}%)")
        nocite = df[df["인용특허명"].astype(str).str.len() == 0]
        chk("인용후보 없는 행은 특허명 인용 없음", True,
            f"{len(nocite)}행 — 질의 유사도 P90 초과 특허가 없어 특허명을 넣지 않았다")

    # ---- 다년차 통합 --------------------------------------------------------
    if "연차수" in df.columns:
        u = df.drop_duplicates("과제고유번호")
        chk("추천에 같은 과제명 중복 없음(연차 통합)",
            u["과제명"].duplicated().sum() == 0,
            f"고유 과제 {len(u)}개 · 중복 과제명 "
            f"{sorted(set(u.loc[u['과제명'].duplicated(), '과제명'].str[:20]))[:3]}"
            if u["과제명"].duplicated().any() else
            f"고유 과제 {len(u)}개 · 통합된 과제 {int(u['연차수'].gt(1).sum())}건"
            f"(최대 {int(u['연차수'].max())}연차)")
        bad_ph = [r for r in u.to_dict("records")
                  if str(r["과제고유번호"]) not in str(r["연차_과제고유번호"]).split(",")]
        chk("대표 과제고유번호가 연차 목록에 포함", not bad_ph, f"위반 {len(bad_ph)}건")
        # 통합 과제의 특허건수는 단일 연차보다 크거나 같아야 한다
        mp = u[u["연차수"] > 1]
        chk("통합 과제 특허건수 ≥ 1", mp.empty or mp["특허건수"].ge(1).all(),
            f"통합 {len(mp)}건 · 특허 중위 {int(mp['특허건수'].median()) if len(mp) else 0}건 "
            f"· 단일 연차 중위 "
            f"{int(u[u['연차수'] == 1]['특허건수'].median()) if (u['연차수'] == 1).any() else 0}건")
    bad_nm = [(r["기업명"], r["과제명"]) for r in df.to_dict("records")
              if bad_project_name(r["과제명"])[0] != bool(r.get("과제명특이", False))]
    chk("과제명 판정 재현 일치", not bad_nm, f"불일치 {len(bad_nm)}건")
    if field and field != "전체":
        bad_f = []
        for r in df.drop_duplicates("과제고유번호").to_dict("records"):
            ok, why = in_medical_domain(str(r.get("과제분야", "")), r["과제명"],
                                        r["과제설명문"])
            if not ok:
                bad_f.append((r["과제명"][:30], why))
        chk(f"추천 과제 전부 의료·제약 연관('{field}')", not bad_f,
            f"위반 {len(bad_f)}건 {bad_f[:3]}" if bad_f
            else f"고유 과제 {df['과제고유번호'].nunique()}개 전부 통과")
    miss_pid = sorted(set(df["과제고유번호"].astype(str)) - set(pat_ok))
    chk("추천 과제가 모두 코퍼스 소속", not miss_pid, f"이탈 {miss_pid[:5]}")
    bad_rank = [n for n, g in df.groupby("기업명")
                if sorted(g["rank"]) != list(range(1, len(g) + 1))]
    chk("기업별 rank 1..n 연속", not bad_rank, f"위반 {bad_rank[:5]}")
    dup = [(n, int(g["과제고유번호"].duplicated().sum())) for n, g in df.groupby("기업명")
           if g["과제고유번호"].duplicated().any()]
    chk("기업별 과제 중복 없음", not dup, f"중복 {dup[:5]}")
    short = df.groupby("기업명").size()
    chk(f"기업별 추천 {final}건", (short == final).all(),
        f"미달 {short[short != final].to_dict()}" if (short != final).any() else f"전부 {final}건")

    # 질의 근거가 요구 우선순위와 일치하는지
    ub = {u["기업명"]: u for u in units}
    wrong = [u["기업명"] for u in units
             if ("수요기술" in u["근거"]) != bool(u["수요기술 내용"])
             or (u["근거"] != ["수요기술"] and not u["기보유기술 내용"])
             or (len(u["근거"]) > 1 and not u["보조사용"])]
    chk("질의 근거 구성 일관", not wrong, f"위반 {wrong[:5]}")
    aug = [f"{u['기업명']}({u['소스']})" for u in units if u["보조사용"]]
    # --verify-only 는 판정을 다시 돌리지 않아 '근거판정' 이력이 없다(근거 구성만 복원한다).
    # 이력이 없으면 이 항목은 건너뛴다 — 없는 키를 참조해 죽지 않게 한다.
    if all("근거판정" in u for u in units):
        chk("보조 근거는 판정 결과와 일치", all(
            u["보조사용"] == any(not ok for _b, ok, _w, _h in u["근거판정"])
            or (u["근거판정"] and u["근거판정"][-1][3] == "force") for u in units),
            f"{len(aug)}개사: " + ", ".join(aug) if aug else "없음")
    else:
        chk("보조 근거 사용 기업", True,
            (f"{len(aug)}개사: " + ", ".join(aug)) if aug else "없음"
            + " (판정 이력 없음 — 저장된 근거 구성으로 확인)")
    chk("기업DB 근거는 사업자번호로 조회",
        all(ub[n]["기업DB_기업명"] for n in
            {u["기업명"] for u in units if "기업DB" in u["근거"]}),
        f"기업DB 근거 {sum(1 for u in units if '기업DB' in u['근거'])}개사")

    # 근거문·금지표현·유망성
    n_r = int((df["판단근거"].astype(str).str.strip() != "").sum())
    n_d = int((df["추천근거_상세"].astype(str).str.strip() != "").sum())
    chk("판단근거 전건 생성", n_r == len(df), f"{n_r}/{len(df)}")
    chk("상세근거 전건 생성", n_d == len(df), f"{n_d}/{len(df)}")
    ban = df[df["판단근거"].astype(str).apply(lambda s: bool(xp.BANNED.search(s)))
             | df["추천근거_상세"].astype(str).apply(lambda s: bool(xp.BANNED.search(s)))]
    chk("근거문에 매칭기준 노출 없음", ban.empty, f"위반 {len(ban)}건")
    txt = "\n".join(df["판단근거"].astype(str)) + "\n".join(df["추천근거_상세"].astype(str))
    chk("근거문에 유망성 점수 노출 없음",
        not re.search(r"유망성\s*(점수|지수)|유망성\s*[:：]?\s*\d", txt))
    neg = df[df["판단근거"].astype(str).str.contains(
        r"부재|미포함|불일치|미흡|부족|없음|한계", regex=True)]
    if not neg.empty:
        warn.append(f"판단근거에 부정 표현 의심 {len(neg)}건: "
                    f"{neg['기업명'].head(3).tolist()}")
    if skipped:
        warn.append(f"질의 근거가 없어 제외된 기업 {len(skipped)}개사: "
                    + ", ".join(f"{s['기업명']}({s['제외사유']})" for s in skipped))
    thin = [f"{u['기업명']}({u['질의길이']}자)" for u in units if u["질의길이"] < QUERY_MIN]
    if thin:
        warn.append(f"질의 본문이 {QUERY_MIN}자 미만 {len(thin)}개사: " + ", ".join(thin))
    if "연계유형" in df.columns:
        from collections import Counter as _CK
        dist = dict(_CK(df["연계유형"]))
        chk("연계유형 전건 판정", df["연계유형"].astype(str).str.strip().ne("").all(), str(dist))
        na = int((df["연계유형"] == KIND_NA).sum())
        if na:
            warn.append(f"연계유형 미분류 {na}건 — LLM 응답에 유형이 없었다")
        bad_k = sorted(set(df["연계유형"]) - set(KINDS) - {KIND_NA})
        chk("연계유형 값이 정의된 범위", not bad_k, f"이상값 {bad_k}")
    cov = df[df["순위구분"] == "기관배분"]
    if not cov.empty:
        warn.append("기관 최소보장으로 넣은 추천 "
                    + ", ".join(f"{r['공급기관']}→{r['기업명']}({r['rank']}위, 적합도 "
                                f"{r['적합도']})" for r in cov.to_dict("records")))
    lowfit = df[df["순위구분"] == "보충"]
    if not lowfit.empty:
        warn.append(f"적합도 하한 미달 '보충' {len(lowfit)}건 "
                    f"({lowfit['기업명'].nunique()}개사) — 후보풀이 얕은 기업")

    log("\n" + "=" * 78 + "\n검증 결과\n" + "=" * 78)
    for ok, name, detail in res:
        log(f"  [{'PASS' if ok else 'FAIL'}] {name:<34} {detail}")
    for w in warn:
        log(f"  [주의] {w}")
    nf = sum(1 for ok, _, _ in res if not ok)
    log(f"\n검증 {len(res)}항목 · 실패 {nf}건 · 주의 {len(warn)}건")
    return nf == 0


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC_XLSX)
    ap.add_argument("--tag", default="MEDITEK_260820")
    ap.add_argument("--only", default="", help="처리할 기업 No(쉼표 구분)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--n-pool", type=int, default=N_POOL)
    ap.add_argument("--n-cand", type=int, default=N_CAND)
    ap.add_argument("--final", type=int, default=FINAL)
    ap.add_argument("--min-fit", type=int, default=MIN_FIT)
    ap.add_argument("--w-cos", type=float, default=W_COS)
    ap.add_argument("--w-prom-pre", type=float, default=W_PROM_PRE)
    ap.add_argument("--w-fit", type=float, default=W_FIT)
    ap.add_argument("--w-cos-re", type=float, default=W_COS_RE)
    ap.add_argument("--w-prom-re", type=float, default=W_PROM_RE)
    ap.add_argument("--w-pat-pre", type=float, default=W_PAT_PRE,
                    help="1차선정 특허축 가중(0=미사용)")
    ap.add_argument("--w-pat-re", type=float, default=W_PAT_RE,
                    help="재랭킹 특허축 가중(0=미사용)")
    ap.add_argument("--no-patent-axis", action="store_true",
                    help="특허 임베딩 축을 쓰지 않는다(이전 점수 체계)")
    ap.add_argument("--patent-emb", default=pax.PAT_EMB, help="특허 임베딩 pkl")
    ap.add_argument("--patsim-tau", type=float, default=pax.TAU_PCT,
                    help="특허 관련 임계값 백분위(질의별로 다시 잡는다)")
    ap.add_argument("--patsim-topk", type=int, default=pax.TOP_K,
                    help="과제당 에너지에 기여하는 특허 수 상한")
    ap.add_argument("--allow-any-name", action="store_true",
                    help="과제명 정상성 검사를 쓰지 않는다(비정상 과제명도 그대로 추천)")
    ap.add_argument("--name-exempt-fit", type=int, default=NAME_EXEMPT_FIT,
                    help="과제명이 비정상이어도 이 적합도 이상이면 예외로 남긴다")
    ap.add_argument("--name-promise-min", type=float, default=NAME_PROMISE_MIN,
                    help="'유망 특허' 판정 기준(등록 특허의 유망성점수 하한)")
    ap.add_argument("--no-merge-phase", action="store_true",
                    help="다년차 과제를 통합하지 않는다(연차별 별개 과제로 취급)")
    ap.add_argument("--cite-patents", type=int, default=CITE_PATENTS,
                    help="근거문에 인용 후보로 넘길 질의 관련 특허 수(0=넘기지 않음)")
    ap.add_argument("--name-extra", type=int, default=NAME_EXTRA,
                    help="과제명 비정상 과제를 후보 정원 밖에서 추가 채점할 건수")
    ap.add_argument("--patent-status", default="생존", choices=("생존", "전체"),
                    help="'생존'=출원 계류·등록 특허만 인정(기본) · '전체'=상태 무관")
    ap.add_argument("--year-min", type=int, default=YEAR_MIN)
    ap.add_argument("--patent-src", default="nice", choices=("nice", "pmeta", "both"),
                    help="특허 1건 이상 판정 소스(nice=특허DB 연계, pmeta=과제메타 특허건수)")
    ap.add_argument("--field", default="의료제약", choices=("의료제약", "전체"),
                    help="매칭 대상 기술 분야 한정(기본: 의료·제약 연관 과제만)")
    ap.add_argument("--allow-subject", action="store_true",
                    help="연구수행주체 필터(대학·출연연 등)도 함께 적용")
    ap.add_argument("--kind-pool", type=int, default=KIND_POOL,
                    help="연계유형을 판정할 상위 후보 수(0=판정 후 가중 없이 표기만)")
    ap.add_argument("--w-kind-co", type=int, default=1,
                    help=f"공동연구 가중 적용(1=적용 {W_KIND_CO}, 0=가중 없음)")
    ap.add_argument("--cover", default=",".join(COVER_SUPPLIERS),
                    help="추천 0건이면 최소 1건을 보장할 공급기관(쉼표 구분). "
                         f"기본값 {','.join(COVER_SUPPLIERS) or '없음'}")
    ap.add_argument("--no-cover", action="store_true", help="최소 보장을 쓰지 않는다")
    ap.add_argument("--dedupe-phase", action="store_true", help="연차 후속과제를 한 건으로 취급")
    ap.add_argument("--no-detail", action="store_true", help="상세근거 생성 생략")
    ap.add_argument("--dry-run", action="store_true", help="LLM 없이 코퍼스·질의 구성만 점검")
    ap.add_argument("--verify-only", action="store_true", help="저장된 산출물만 재검증")
    ap.add_argument("--augment", default="",
                    help="판정을 무시하고 보조 근거를 강제로 더할 기업 No(쉼표 구분)")
    ap.add_argument("--no-augment", action="store_true",
                    help="보조 근거를 쓰지 않는다(근거 충분성 판정 끄기)")
    ap.add_argument("--judge-only", action="store_true",
                    help="근거 충분성 판정만 수행하고 결과 표를 출력한다(매칭 안 함)")
    ap.add_argument("--relabel", action="store_true",
                    help="LLM 없이 표시용 기술명만 다시 계산해 pkl/xlsx/JSON 갱신")
    a = ap.parse_args()

    pkl = os.path.join(cm.OUT_DIR, f"{a.tag}_매칭.pkl")
    xlsx = os.path.join(cm.OUT_DIR, f"{a.tag}_매칭.xlsx")
    jsn = os.path.join(cm.OUT_DIR, f"{a.tag}_보고서.json")

    log(f"· 기업 pkl 로드(사업자번호 색인)… ({COMPANY_PKL})")
    cidx = company_index()
    log(f"  사업자번호 {len(cidx)}건")
    force = {x.strip() for x in a.augment.split(",") if x.strip()}
    units_all, skipped = load_units(a.src, cidx)
    units = units_all
    sup = [_s(x) for x in pd.read_excel(a.src, sheet_name=SH_SUPPLY)["기관명"] if _s(x)]

    def _restore_bases():
        """저장된 근거 구성(supply_<tag>_bases.json)을 units 에 되살린다.

        --relabel·--verify-only 는 판정을 다시 돌리지 않으므로, 이걸 하지 않으면 단위가
        우선순위 첫 근거만 가진 상태가 되어 저장된 산출물과 어긋난다.
        """
        p = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_bases.json")
        if not os.path.exists(p):
            return
        saved = json.load(open(p, encoding="utf-8"))
        for u in units:
            src = saved.get(u["키"])
            if src:
                apply_bases(u, src.split("+"))

    if a.relabel:
        _restore_bases()
        df = pd.read_pickle(pkl)
        kw_ck = json.load(open(os.path.join(cm.OUT_DIR,
                                            f"supply_{a.tag}_keywords_ckpt.json"),
                               encoding="utf-8"))
        keep = set(df["기업명"])
        units = [u for u in units if u["기업명"] in keep]
        for u in units:
            u["키워드"] = kw_ck.get(u["키"], [])
        new = {u["기업명"]: u["수요기술명"] for u in units}
        chg = [(n, o, new[n]) for n, o in
               df.drop_duplicates("기업명").set_index("기업명")["수요기술명"].items()
               if new.get(n, o) != o]
        df["수요기술명"] = df["기업명"].map(lambda n: new.get(n, ""))
        df.to_pickle(pkl)
        save_xlsx(df, xlsx)
        report_json(df, units, jsn)
        log(f"표시용 기술명 재계산 {len(chg)}개사 변경 → {pkl} / {xlsx} / {jsn}")
        for n, o, x in chg:
            log(f"    {n:<18} {o!r}\n{'':<22}→ {x!r}")
        c = build_corpus(sup, a.year_min, a.patent_src, a.allow_subject, a.field,
                         a.patent_status, not a.allow_any_name, a.name_promise_min,
                         not a.no_merge_phase)
        raise SystemExit(0 if verify(df, units, skipped, c, a.final, a.patent_src, a.field,
                          a.patent_status) else 1)

    if a.verify_only:
        _restore_bases()
        df = pd.read_pickle(pkl)
        c = build_corpus(sup, a.year_min, a.patent_src, a.allow_subject, a.field,
                         a.patent_status, not a.allow_any_name, a.name_promise_min,
                         not a.no_merge_phase)
        ok = verify(df, [u for u in units if u["기업명"] in set(df["기업명"])],
                    skipped, c, a.final, a.patent_src, a.field, a.patent_status)
        raise SystemExit(0 if ok else 1)

    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        units = [u for u in units if u["번호"] in keep]
    if a.limit:
        units = units[:a.limit]

    kw_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_keywords_ckpt.json")
    sc_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_scores_ckpt.json")
    ex_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_explain_v5_ckpt.json")
    rs_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_reason_v5_ckpt.json")
    kd_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_kind_ckpt.json")
    kd_ck = json.load(open(kd_path, encoding="utf-8")) if os.path.exists(kd_path) else {}
    ck = {p: (json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {})
          for p in (kw_path, sc_path, ex_path, rs_path)}
    kw_ck, sc_ck, ex_ck, rs_ck = (ck[kw_path], ck[sc_path], ck[ex_path], ck[rs_path])
    if any(ck.values()):
        log(f"체크포인트: 키워드 {len(kw_ck)} · 점수 {len(sc_ck)} · "
            f"상세 {len(ex_ck)} · 판단 {len(rs_ck)}")
    # ---- 질의 근거 확정(35B 판정) ----
    jd_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_judge_ckpt.json")
    jd_ck = json.load(open(jd_path, encoding="utf-8")) if os.path.exists(jd_path) else {}
    if not a.dry_run:
        log("· 35B 로드…")
        cm.load_model_blocking(progress_cb=lambda m: log("  " + m))
    log(f"\n· 질의 근거 충분성 판정({'35B' if not a.dry_run else '길이 규칙'})… "
        f"판정 캐시 {len(jd_ck)}건")
    select_bases(units_all, jd_ck, force=force, disable=a.no_augment, log_fn=log,
                 use_llm=not a.dry_run)
    json.dump(jd_ck, open(jd_path, "w", encoding="utf-8"), ensure_ascii=False)
    from collections import Counter as _C
    log(f"  확정 근거 분포 {dict(_C(u['소스'] for u in units_all))}")
    tgt = [u for u in units_all if u["보조사용"]]
    if tgt:
        log("  보조 근거 사용 " + str(len(tgt)) + "개사: "
            + ", ".join(f"[{u['번호']}] {u['기업명']}({u['소스']})" for u in tgt))
    miss = force - {u["번호"] for u in tgt} - {u["키"] for u in tgt}
    if miss:
        log(f"! 강제 지정했으나 쓸 보조 근거가 없는 기업: {sorted(miss)}")

    # 근거 구성이 바뀐 기업은 옛 키워드·점수·근거를 재사용하면 안 된다.
    # 상태 파일에는 '저장된 산출물이 어떤 근거로 만들어졌는지'가 들어 있다. 기록이 없는
    # 기업도 '바뀐 것'으로 본다 — 모르는 채 옛 점수를 쓰는 쪽이 훨씬 위험하다.
    st_path = os.path.join(cm.OUT_DIR, f"supply_{a.tag}_bases.json")
    prev = json.load(open(st_path, encoding="utf-8")) if os.path.exists(st_path) else {}
    changed = [u for u in units if prev.get(u["키"]) != u["소스"]]   # 매칭 대상만
    for u in changed:
        kw_ck.pop(u["키"], None)
        sc_ck.pop(u["키"], None)
        pre = f"{u['번호'] or u['키']}::"
        for d in (ex_ck, rs_ck):
            for k in [k for k in d if k.startswith(pre)]:
                d.pop(k)
        for k in [k for k in kd_ck if k.startswith(f"{u['키']}::")]:
            kd_ck.pop(k)
    if changed:
        for path, d in ((kw_path, kw_ck), (sc_path, sc_ck),
                        (ex_path, ex_ck), (rs_path, rs_ck)):
            json.dump(d, open(path, "w", encoding="utf-8"), ensure_ascii=False)
        log(f"  근거 변경 {len(changed)}개사 → 체크포인트 폐기: "
            + ", ".join(f"{u['기업명']}({prev.get(u['키']) or '기록없음'}→{u['소스']})"
                        for u in changed))

    def _save_bases():
        """저장된 산출물의 근거 구성을 기록(병합 저장 — --only 로 일부만 돌려도 안 지워진다)."""
        cur = json.load(open(st_path, encoding="utf-8")) if os.path.exists(st_path) else {}
        cur.update({u["키"]: u["소스"] for u in units})
        json.dump(cur, open(st_path, "w", encoding="utf-8"), ensure_ascii=False)

    if a.judge_only:
        log("\n" + "=" * 78 + "\n근거 충분성 판정 결과\n" + "=" * 78)
        log(f"  {'No':>4} {'기업명':<18}{'확정 근거':<20}{'질의':>7}  판정 이력")
        for u in units_all:
            hist = " | ".join(f"{b}:{'충분' if ok else '불충분'}({how})"
                              for b, ok, _why, how in u["근거판정"]) or "(추가 근거 없음)"
            log(f"  {u['번호']:>4} {u['기업명']:<18}{u['소스']:<20}{u['질의길이']:>6}자  {hist}")
        log(f"\n보조 근거 사용 {len(tgt)}/{len(units_all)}개사")
        return

    from collections import Counter
    log(f"\n대상 기업 {len(units)}개사 · 질의 근거 {dict(Counter(u['소스'] for u in units))}")
    for u in units:
        log(f"    [{u['번호']:>3}] {u['기업명']:<16} {u['소스']:<7} 질의 {u['질의길이']:>5}자"
            + ("  (기보유기술 보유 — 요구 우선순위상 질의에는 미사용)"
               if u["기보유기술_보유"] and u["소스"] == "수요기술" else "")
            + (f"  ← 기업DB '{u['기업DB_기업명']}'" if u["소스"] == "기업DB" else ""))
    if skipped:
        log(f"제외 {len(skipped)}개사: "
            + " · ".join(f"{s['기업명']}({s['제외사유']})" for s in skipped))

    c = build_corpus(sup, a.year_min, a.patent_src, a.allow_subject, a.field,
                         a.patent_status, not a.allow_any_name, a.name_promise_min,
                         not a.no_merge_phase)
    rev = org2supply(c["resolved"])

    log("\n· pro-sroberta 로드…")
    model = SentenceTransformer(cm.MODEL_DIR, device="cpu")
    model.eval()
    encode = cm.make_encoder(model)


    log("\n· 키워드 추출…")
    for n, u in enumerate(units, 1):
        if u["키"] not in kw_ck:
            if a.dry_run:                        # LLM 없이도 규칙기반 폴백으로 확인 가능
                kw_ck[u["키"]] = mt.filter_keywords(
                    cm.extract_keywords_doc(query_doc(u)))
            else:
                kw_ck[u["키"]] = keywords_of(u)
            json.dump(kw_ck, open(kw_path, "w", encoding="utf-8"), ensure_ascii=False)
        u["키워드"] = kw_ck[u["키"]]
        log(f"    [{u['번호']:>3}] {u['기업명']:<16} 키워드 {len(u['키워드']):>2}개: "
            f"{', '.join(u['키워드'][:8])} [{n}/{len(units)}]")

    if a.no_patent_axis:
        a.w_pat_pre = a.w_pat_re = 0.0
    axis = None
    if a.w_pat_pre > 0 or a.w_pat_re > 0:
        pax.TAU_PCT = a.patsim_tau
        pax.TOP_K = a.patsim_topk
        axis = pax.load(c["pid"], src=a.patent_emb, log=log,
                        members=c.get("members"))

    wsum = a.w_fit + a.w_cos_re + a.w_prom_re + a.w_pat_re
    psum = a.w_cos + a.w_prom_pre + a.w_pat_pre
    log(f"\n· 1차선정 가중치 유사도={a.w_cos} 유망성={a.w_prom_pre} "
        f"특허={a.w_pat_pre} → 후보 {a.n_cand}건 (풀 {a.n_pool})")
    log(f"· 재랭킹 가중치 적합도={a.w_fit} 유사도={a.w_cos_re} 유망성={a.w_prom_re} "
        f"특허={a.w_pat_re} · 적합도 하한 {a.min_fit} → TOP{a.final}")

    prom_all = np.clip(c["promise"] / 100.0, 0, 1)
    cover_set = set() if a.no_cover else {x.strip() for x in a.cover.split(",") if x.strip()}
    unknown_cover = cover_set - set(so.SUPPLY_ORGS)
    if unknown_cover:
        raise SystemExit(f"[중단] 대응표에 없는 최소보장 공급기관: {sorted(unknown_cover)}")
    cover_pool = {}
    rows, t0 = [], time.time()
    for n, u in enumerate(units, 1):
        qv = encode(cm.SEP.join(u["키워드"]))
        cos = np.clip(c["M"] @ qv, -1, 1)
        if axis is not None:
            pat_all, pd_ = axis.score(qv)
        else:
            pat_all = np.zeros(len(cos))
            pd_ = {"tau": 0.0, "cnt": np.zeros(len(cos)), "E": np.zeros(len(cos)),
                   "e_n": np.zeros(len(cos)), "legal": np.zeros(len(cos)),
                   "recent": np.zeros(len(cos)), "max": np.zeros(len(cos)),
                   "legal_max": np.zeros(len(cos))}
        pool = np.argsort(-cos)[:min(a.n_pool, len(cos))]
        lo, hi = float(cos[pool].min()), float(cos[pool].max())
        cos_n = (cos - lo) / (hi - lo) if hi > lo else np.zeros_like(cos)
        pre = np.full(len(cos), -1.0)
        pre[pool] = (a.w_cos * np.clip(cos_n[pool], 0, 1)
                     + a.w_prom_pre * prom_all[pool]
                     + a.w_pat_pre * pat_all[pool]) / psum
        if axis is not None:
            log(f"      [{u['번호']:>3}] {u['기업명']:<16} 특허축 τ={pd_['tau']:.3f} "
                f"· 관련특허 보유과제 {int((pd_['cnt'] > 0).sum())}건 "
                f"· 특허축 상위10 평균 {float(np.sort(pat_all)[-10:].mean()):.3f}")
        # 과제명이 비정상인 과제는 **후보 정원을 잡아먹지 않게** 별도로 덧붙인다.
        # 예외 판정에 적합도가 필요해 채점은 해야 하는데, 이들이 정원 안에 들어오면
        # (특허 포트폴리오가 커서 1차점수가 높다) 정상 과제를 밀어낸다 —
        # 실측에서 에이아이다이콤 적합도 95 짜리 1위 과제가 후보에서 탈락했다.
        bad_n = c["namebad"] if c.get("check_name") else np.zeros(len(cos), bool)
        ord_pool = pool[np.argsort(-pre[pool])]
        cand_idx = ord_pool[~bad_n[ord_pool]][:a.n_cand]
        if a.name_extra > 0:
            extra = ord_pool[bad_n[ord_pool]][:a.name_extra]
            if len(extra):
                cand_idx = np.concatenate([cand_idx, extra])
        cands = [{"id": j + 1, "과제명": c["pname"][i], "과제설명문": c["pdesc"][i],
                  "키워드": list(c["pkw"][i]) if isinstance(c["pkw"][i], (list, tuple))
                  else []}
                 for j, i in enumerate(cand_idx)]

        got = sc_ck.get(u["키"], {})
        need = [x for x in cands if str(c["pid"][cand_idx[x["id"] - 1]]) not in got]
        if need and not a.dry_run:
            log(f"    [{u['번호']:>3}] {u['기업명']}: LLM 재랭킹 채점 "
                f"{len(need)}/{len(cands)}건…")
            sc = score_candidates(u, need)
            for x in need:
                got[str(c["pid"][cand_idx[x["id"] - 1]])] = list(sc[x["id"]])
            sc_ck[u["키"]] = got
            json.dump(sc_ck, open(sc_path, "w", encoding="utf-8"), ensure_ascii=False)

        scored = []
        for i in cand_idx:
            rec = got.get(str(c["pid"][i]), (0, ""))
            fit, reason = rec[0], rec[-1]
            total = 100.0 * (a.w_fit * fit / 100.0 + a.w_cos_re * float(cos_n[i])
                             + a.w_prom_re * float(prom_all[i])
                             + a.w_pat_re * float(pat_all[i])) / wsum
            scored.append({"i": int(i), "fit": int(fit), "reason": reason,
                           "total": total, "pre": float(pre[i]),
                           "cos": float(cos[i]), "cos_n": float(cos_n[i]),
                           "pat": float(pat_all[i]),
                           "pat_e": float(pd_["E"][i]), "pat_cnt": int(pd_["cnt"][i]),
                           "pat_sim_e": float(pd_["e_n"][i]),
                           "pat_legal": float(pd_["legal"][i]),
                           "pat_recent": float(pd_["recent"][i]),
                           "pat_max": float(pd_["max"][i]), "pat_tau": float(pd_["tau"]),
                           "pat_legal_max": float(pd_["legal_max"][i])})
        scored.sort(key=lambda x: (-x["total"], -x["fit"], -x["cos"]))

        # ---- 유형 판정 후 가중 적용해 재정렬(상위 후보만 판정한다) ----
        if not a.dry_run and a.kind_pool > 0:
            pool = scored[:a.kind_pool]
            kmap = classify_kinds(u, pool, c, kd_ck)
            json.dump(kd_ck, open(kd_path, "w", encoding="utf-8"), ensure_ascii=False)
            for x in scored:
                kind = kmap.get(str(c["pid"][x["i"]]), (KIND_NA, ""))[0]
                x["kind"] = kind
                w_k = W_KIND_CO if (kind == KIND_CO and a.w_kind_co) else 1.0
                if w_k != 1.0:
                    i = x["i"]
                    x["total"] = 100.0 * (a.w_fit * (x["fit"] / 100.0) * w_k
                                          + a.w_cos_re * float(cos_n[i])
                                          + a.w_prom_re * float(prom_all[i])
                                          + a.w_pat_re * float(pat_all[i])) / wsum
            scored.sort(key=lambda x: (-x["total"], -x["fit"], -x["cos"]))
            n_co = sum(1 for x in pool if x.get("kind") == KIND_CO)
            log(f"      [{u['번호']:>3}] {u['기업명']:<16} 유형 판정 상위 {len(pool)}건 "
                f"→ 기술도입 {len(pool) - n_co} · 공동연구 {n_co} (가중 {W_KIND_CO} 적용)")

        def name_ok(x):
            """과제명이 연구 주제가 아닌 후보는 예외 조건을 충족할 때만 통과시킨다."""
            i = x["i"]
            if not c.get("check_name") or not c["namebad"][i]:
                return True
            if c["namewhy"][i] == "기관명 단독":
                return False          # 적합도·특허와 무관하게 예외 없음
            return (x["fit"] >= a.name_exempt_fit
                    and int(c["pat_promise"][i]) >= 1)

        n_drop, drop_ex = 0, []
        seen, sel = set(), []
        for pool_pass in (True, False):           # 적합도 하한 통과분 우선, 부족분만 보충
            for x in scored:
                if (x["fit"] >= a.min_fit) != pool_pass:
                    continue
                if not name_ok(x):
                    n_drop += 1
                    if len(drop_ex) < 3:
                        drop_ex.append(f"{c['pname'][x['i']][:22]}"
                                       f"({c['namewhy'][x['i']]}·적합도 {x['fit']}"
                                       f"·유망특허 {int(c['pat_promise'][x['i']])})")
                    continue
                tk = cm.title_key(c["pname"][x["i"]], a.dedupe_phase)
                if tk in seen:
                    continue
                seen.add(tk)
                sel.append(dict(x, 순위구분="적합" if pool_pass else "보충"))
                if len(sel) >= a.final:
                    break
            if len(sel) >= a.final:
                break
        if axis is not None:                 # 근거문에 인용할 질의 관련 특허명(상위 3건)
            for x in sel:
                x["pat_titles"] = axis.relevant(qv, x["i"], n=a.cite_patents)
        n_pass = sum(1 for x in scored if x["fit"] >= a.min_fit)
        if n_drop:
            log(f"      [{u['번호']:>3}] {u['기업명']:<16} 과제명 비정상 후보 {n_drop}건 제외 "
                f"(예외: 적합도 ≥{a.name_exempt_fit} ∧ 유망 등록특허 ≥1) — "
                + ", ".join(drop_ex))
        n_ex = sum(1 for x in sel if c["namebad"][x["i"]])
        if n_ex:
            log(f"      [{u['번호']:>3}] {u['기업명']:<16} 과제명 비정상이나 예외로 유지 "
                f"{n_ex}건: " + ", ".join(
                    f"{c['pname'][x['i']][:22]}(적합도 {x['fit']}·유망특허 "
                    f"{int(c['pat_promise'][x['i']])})"
                    for x in sel if c["namebad"][x["i"]]))

        # pool 밖에서 올라온 후보가 있으면 여기서 보완 판정한다(대개 캐시에 이미 있다)
        kinds = ({} if a.dry_run else
                 classify_kinds(u, sel, c, kd_ck))
        if not a.dry_run:
            json.dump(kd_ck, open(kd_path, "w", encoding="utf-8"), ensure_ascii=False)
        for rank, x in enumerate(sel, 1):
            kind, why = kinds.get(str(c["pid"][x["i"]]), (KIND_NA, ""))
            rows.append(make_row(u, x, rank, c, rev, ex_ck, rs_ck, ex_path, rs_path,
                                 a.dry_run, a.no_detail, kind, why))
        for x in scored:                 # 최소 보장용 후보 적립(공급기관별 최고 1건)
            sup = rev.get(c["org"][x["i"]], "")
            if sup in cover_set and str(c["pid"][x["i"]]) not in {
                    str(c["pid"][y["i"]]) for y in sel}:
                cover_pool.setdefault(sup, []).append((x, u))
        best = sel[0] if sel else None
        el = time.time() - t0
        log(f"    [{u['번호']:>3}] {u['기업명']:<16} 후보 {len(cands)} → 하한통과 {n_pass}"
            f" → TOP{len(sel)}"
            + (f" (1위 {best['total']:.1f}점 적합도 {best['fit']} "
               f"특허 {int(c['pat'][best['i']])}건 {c['org'][best['i']]})" if best else "")
            + f" [{n}/{len(units)}] {el/60:.1f}분")

    # ---- 최소 보장: 추천 0건인 공급기관에 최고 후보 1건을 배분 ----
    if cover_set and len(units) < len(units_all):
        log(f"\n· 최소 보장 건너뜀 — 일부 기업만 돌린 실행에서는 적용하지 않는다 "
            f"(대상 {sorted(cover_set)})")
    elif cover_set:
        have = {r["공급기관"] for r in rows}
        for sup in sorted(cover_set):
            if sup in have:
                log(f"\n· 최소 보장 불필요 — {sup} 이미 "
                    f"{sum(1 for r in rows if r['공급기관'] == sup)}건")
                continue
            pool = sorted(cover_pool.get(sup, []), key=lambda t: (-t[0]["fit"],
                                                                 -t[0]["total"]))
            if not pool:
                log(f"\n! 최소 보장 실패 — {sup}: 채점된 후보가 없다")
                continue
            x, u = pool[0]
            tgt = [r for r in rows if r["기업명"] == u["기업명"]]
            drop = max(tgt, key=lambda r: r["rank"])          # 가장 낮은 순위를 교체
            kd = classify_kinds(u, [x], c, kd_ck) if not a.dry_run else {}
            json.dump(kd_ck, open(kd_path, "w", encoding="utf-8"), ensure_ascii=False)
            k2, w2 = kd.get(str(c["pid"][x["i"]]), (KIND_NA, ""))
            rows[rows.index(drop)] = make_row(
                u, dict(x, 순위구분="기관배분"), drop["rank"], c, rev,
                ex_ck, rs_ck, ex_path, rs_path, a.dry_run, a.no_detail, k2, w2)
            log(f"\n· 최소 보장 배분 — {sup}: [{u['번호']}] {u['기업명']} "
                f"{drop['rank']}위 교체 (적합도 {drop['적합도']}→{x['fit']}, "
                f"{drop['과제명'][:28]}… → {c['pname'][x['i']][:28]}…)")

    df = pd.DataFrame(rows, columns=OUT_COLS)
    if a.dry_run:
        log(f"\n[dry-run] {len(df)}행 구성 확인 (저장 안 함)")
        verify(df, units, skipped, c, a.final, a.patent_src, a.field,
                          a.patent_status)
        return

    partial = len(units) < len(units_all)
    if partial and os.path.exists(pkl):
        # 일부 기업만 다시 돌린 경우 — 나머지 기업의 기존 추천은 그대로 두고 갈아끼운다
        old_df = pd.read_pickle(pkl)
        import gen_report as _g
        for col in OUT_COLS:         # 컬럼이 늘었으면(예: 과제분야) 채워 넣는다
            if col not in old_df.columns:
                old_df[col] = ([_g.extract_class(str(d)) or "" for d in old_df["과제설명문"]]
                               if col == "과제분야" else "")
        redone = {u["기업명"] for u in units}
        kept = old_df[~old_df["기업명"].isin(redone)][OUT_COLS]
        df = pd.concat([kept, df], ignore_index=True)
        order = {u["기업명"]: i for i, u in enumerate(units_all)}
        df = (df.assign(_o=df["기업명"].map(order))
                .sort_values(["_o", "rank"]).drop(columns="_o").reset_index(drop=True))
        log(f"  부분 재매칭 병합: 유지 {len(kept)}행({old_df['기업명'].nunique() - len(redone)}개사)"
            f" + 갱신 {len(rows)}행({len(units)}개사) → {len(df)}행")
    df.to_pickle(pkl)
    save_xlsx(df, xlsx)
    report_json(df, units_all, jsn)
    # 연차 통합 대응표를 남긴다 — 보고서 특허 표(_patent_prep_nice.py)가 같은 기준으로
    # 전 연차 특허를 합쳐야 표의 건수와 매칭 결과의 특허건수가 어긋나지 않는다.
    ph_path = os.path.join(cm.OUT_DIR, f"{a.tag}_연차통합.json")
    pmg.save(c.get("members", {}), ph_path)
    log(f"저장: {ph_path} (연차 통합 {len(c.get('members', {}))}개 그룹)")
    _save_bases()                    # 산출물과 근거 구성 기록을 함께 갱신(매칭한 기업만)
    log(f"\n✔ 저장: {pkl} / {xlsx} / {jsn} "
        f"({len(units)}개사 · {len(df)}행 · {(time.time()-t0)/60:.1f}분)")
    log(f"모델: {cm.MODEL_ID} · 근거 생성 계측 {xp.STATS}")
    verify(df, [u for u in units_all if u["기업명"] in set(df["기업명"])],
           skipped, c, a.final, a.patent_src, a.field)


if __name__ == "__main__":
    main()
