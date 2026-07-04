# -*- coding: utf-8 -*-
"""
user_input_keywords.py

matching_viewer 의 '사용자 자유입력 → 명사 키워드 정제' 로직을 분리한 모듈.

배경
----
저장된 임베딩(company_embeddings / public_RnD_embeddings)은 모두 '키워드 리스트'
즉 짧은 명사 키워드 문체로 생성되어 있다. 그런데 사용자가 유사 검색 시 추가로
입력하는 값(주요 제품 / 사업목적 / 한글 키워드 / 과학기술표준분류)은 사업자등록
업종 문장체("반도체 제품 도매업", "전기전자 제품 제조 판매업")인 경우가 많다.
이 문장체를 그대로 임베딩하면 행정/정책 문체 영역으로 끌려가 엉뚱한 과제·기업과
매칭되므로, 핵심 명사만 추출해 기존 키워드 리스트와 같은 문체로 맞춰준다.

불용어(USER_INPUT_STOPWORDS)와 업종 접미사(USER_INPUT_SUFFIXES)는
gen_compound_keywords.py 의 STOPWORDS / JOSA_EOMI 와 동일 어휘를 재사용한다.

공개 API
--------
- normalize_user_input(text)            : 자유입력 → 정제된 명사 키워드 리스트
- is_meaningless_code(token)            : 과기표준분류 코드 등 무의미 토큰 판정
- build_combined_company_text(row, ud)  : 기업 행 + 사용자입력 → 임베딩 입력 문자열
- build_combined_project_text(row, ud)  : 과제 행 + 사용자입력 → 임베딩 입력 문자열
"""

import re

import pandas as pd


# 키워드 구분자 — 임베딩 입력 문자열을 ';' 로 join
KEYWORD_SEP = ";"

# 사용자 자유입력(주요 제품/사업목적 등) 정제용 불용어.
# 업종 문장체에 흔한 일반 명사·동사·조사를 제거해 핵심 명사만 남긴다.
USER_INPUT_STOPWORDS = {
    "기타", "제조업", "제품", "사업", "회사", "업체", "관련", "서비스",
    "제조", "판매", "생산", "개발", "기술", "산업", "업무", "운영",
    "관리", "지원", "제공", "이용", "사용", "활용", "처리", "수행",
    "분야", "부문", "종류", "형태", "방식", "방법", "과정", "절차",
    "대상", "범위", "내용", "항목", "종목", "품목", "물품", "물자",
    "시설", "설비", "장비", "기기", "기계", "장치", "도구", "용품",
    "자재", "재료", "원료", "부품", "소재", "물질", "성분", "요소",
    "구조", "형식", "유형", "종별", "구분", "분류", "목록",
    "도매업", "부대", "일체", "일반", "임대업", "판매업", "상기",
    "각호", "공급", "호에", "서비스업", "소매업", "부동산", "제작",
    "매매", "전문", "응용", "목적", "대행업", "신품", "공업",
    "소매", "가공", "상거래", "도매", "형성", "경영", "유지",
    "기자재", "단계", "용역", "특수", "작업", "유사", "조립", "대행",
    "가공업", "무역업", "위", "각항", "부대되는", "사업일체",
    "실시", "진행", "추진", "시행", "완료", "종료", "시작", "착수",
    "설립", "설치", "구축", "도입", "적용", "반영", "포함", "제외",
    "및", "외", "내", "중", "상", "하", "전", "후", "등", "것", "수",
}

# 업종 접미사 — 제거 후 남는 핵심 명사만 키워드로 사용
# 예) "반도체도매업" → "반도체", "정밀화학공업" → "정밀화학"
USER_INPUT_SUFFIXES = (
    "판매업", "도매업", "소매업", "서비스업", "제조업", "생산업",
    "가공업", "임대업", "대행업", "무역업", "유통업", "공업", "업",
)

# 어절 끝에 붙은 조사(助詞) — 제거 후 남는 핵심 명사만 키워드로 사용.
# 예) "생명공학에" → "생명공학", "원료의약품의" → "원료의약품".
# 길이 긴 것부터 매칭하며, 제거 후 남는 어간이 2글자 미만이면 적용하지 않는다.
# gen_compound_keywords.JOSA_EOMI 의 조사 어휘와 정합되게 구성하되,
# 한 글자 조사 중 명사 말음과 충돌이 잦은 것(가/이/은/는/도/만/로/와/과)은
# 오제거(예: 결과→결, 참가→참)를 피하려 의도적으로 제외한다.
USER_INPUT_JOSA = (
    "으로서", "으로써", "에서의", "에게서", "에서", "에게", "으로",
    "와의", "과의", "에의", "에", "의", "을", "를",
)

# 연결어미·용언 활용·접속어 — 자유입력 문장에 섞여 들어온 비명사 어절.
# 예) "생명공학에 관련된 시약" 의 "관련된", "~을 위한", "~에 대한" 등.
# gen_compound_keywords.JOSA_EOMI 의 어미/용언·접속어 어휘를 재사용한다.
USER_INPUT_CONNECTIVES = {
    "관련된", "관련되", "위한", "통한", "대한", "관한", "따른", "인한",
    "위해", "통해", "위하여", "하는", "되는", "있는", "없는", "같은",
    "하고", "되고", "하며", "되며", "해서", "하여", "되어", "하면", "되면",
    "또는", "혹은", "그리고", "하지만", "그러나", "따라서", "그래서",
    "해당하는", "필요한", "가능한", "부대하는", "속하는", "부대되는",
}

# 불용어 동사의 활용형 어미 — '{불용어}+어미'(예: 이용한·활용하여·개발하는)를 제거하기 위함.
# 어간이 USER_INPUT_STOPWORDS 에 있을 때만 토큰을 버리므로 일반 명사는 손상되지 않는다.
# (긴 어미부터 매칭)
USER_INPUT_VERB_EOMI = ("하여", "하는", "되어", "되는", "한", "해", "된", "되", "함", "됨")


def is_meaningless_code(token):
    """과기표준분류 코드('EH030301')처럼 의미 없는 영숫자 코드·순수 숫자 토큰인지 판정.

    한글이 없는 ASCII 토큰만 대상으로 한다(한글 키워드는 항상 통과):
    - 순수 숫자('030301', '2024')는 길이 무관 제거
    - 영문+숫자 혼합('EH030301', 'RS232')은 5글자 이상일 때 코드로 보고 제거
    길이 4 이하의 혼합 토큰('5G', '3D', '4K', 'MP3', 'H264')은 의미 있는 기술 용어로
    간주해 보존한다. 순수 영문 약어('AI', 'IoT', 'RFID')는 숫자가 없어 항상 보존된다.
    """
    if re.search(r'[가-힣]', token):
        return False
    if token.isdigit():
        return True
    if len(token) >= 5 and re.fullmatch(r'[A-Za-z0-9]+', token) \
            and re.search(r'[A-Za-z]', token) and re.search(r'\d', token):
        return True
    return False


def normalize_user_input(text):
    """사용자 자유입력(주요 제품/사업목적 등)을 저장 임베딩과 같은 '명사 키워드' 문체로 정제.

    업종 문장체("반도체 제품 도매업", "전기전자 제품 제조 판매업")를 그대로 임베딩하면
    행정/정책 문체 영역으로 끌려가 엉뚱한 과제와 매칭되므로, 핵심 명사만 추출한다.
    절차:
    1) 구두점/구분자를 공백으로 치환 후 어절 단위로 분리
    2) 2글자 미만·불용어·연결어미(관련된/위한/대한 등) 제거
    3) 의미 없는 코드(과기표준분류 코드 'EH030301' 등 영숫자 코드·순수 숫자) 제거
    4) 업종 접미사(~판매업/~도매업/~업 등) 제거
    5) 어절 끝 조사(~에/~의/~을/~으로 등) 제거 후 남는 핵심 명사 재검사
    반환: 정제된 키워드 토큰 리스트(등장 순서 유지).
    """
    if not text:
        return []
    text = re.sub(r'[^\w\s가-힣a-zA-Z0-9]', ' ', str(text))
    out = []
    for w in text.split():
        wc = w.casefold()
        if len(w) < 2 or wc in USER_INPUT_STOPWORDS or wc in USER_INPUT_CONNECTIVES:
            continue
        if is_meaningless_code(w):
            continue
        stem = w
        for suf in USER_INPUT_SUFFIXES:
            if stem.endswith(suf) and len(stem) > len(suf):
                stem = stem[:-len(suf)]
                break
        # 어절 끝 조사 제거 (제거 후 어간이 2글자 이상으로 남을 때만)
        for josa in USER_INPUT_JOSA:
            if stem.endswith(josa) and len(stem) - len(josa) >= 2:
                stem = stem[:-len(josa)]
                break
        # 불용어 동사 활용형 제거: 어미를 뗀 어간이 불용어이면 토큰 전체를 버린다.
        # (예: 이용한→'이용', 활용하여→'활용', 개발하는→'개발' 이 모두 불용어)
        drop_verb = False
        for em in USER_INPUT_VERB_EOMI:
            if stem.endswith(em) and stem[:-len(em)].casefold() in USER_INPUT_STOPWORDS:
                drop_verb = True
                break
        if drop_verb:
            continue
        stem = stem.strip()
        sc = stem.casefold()
        if len(stem) < 2 or sc in USER_INPUT_STOPWORDS or sc in USER_INPUT_CONNECTIVES:
            continue
        out.append(stem)
    return out


def _join_combined(row, user_texts, fallback_col):
    """키워드_리스트 + 사용자 입력 명사 키워드 + (조건부) fallback 컬럼을 ';' 로 묶는다.

    - 기존 키워드 리스트와 사용자 입력 토큰이 겹치면 한 번만 사용한다.
      비교는 strip + casefold 정규화 키로 수행하되, 출력에는 먼저 등장한 원형을 유지한다.
    - fallback_col(기업=10차산업코드명, 과제=과제명)은 기존 키워드_리스트 길이가
      3 이하일 때만 보조로 덧붙이되, 문장 통째가 아니라 normalize_user_input 으로
      명사 키워드만 추출(불용어·연결어·조사·업종접미사 정제)해 덧붙인다.

    row: pandas Series 또는 .get 을 지원하는 매핑
    user_texts: 사용자 자유입력 문자열들 (예: [주요제품, 사업목적])
    fallback_col: 키워드가 빈약할 때 덧붙일 컬럼명
    """
    parts = []
    seen = set()

    def add(token):
        if not token:
            return
        key = token.strip().casefold()
        if not key or key in seen:
            return
        seen.add(key)
        parts.append(token.strip())

    kw_list = row.get('키워드_리스트') if hasattr(row, 'get') else None
    kw_list_len = len(kw_list) if isinstance(kw_list, (list, tuple)) else 0
    if isinstance(kw_list, (list, tuple)):
        for k in kw_list:
            if k is None:
                continue
            add(k if isinstance(k, str) else str(k))

    for text in user_texts:
        if not text:
            continue
        for tok in normalize_user_input(text):
            add(tok)

    if kw_list_len <= 3:
        fb = row.get(fallback_col) if hasattr(row, 'get') else None
        if fb is not None:
            try:
                fb_na = bool(pd.isna(fb))
            except (TypeError, ValueError):
                fb_na = False
            if not fb_na:
                # 과제명/10차산업코드명도 문장 통째가 아니라 normalize_user_input 으로
                # 명사 키워드만 추출(불용어·연결어·조사·업종접미사 정제)해 추가한다.
                for tok in normalize_user_input(str(fb)):
                    add(tok)

    return KEYWORD_SEP.join(parts)


def build_combined_company_text(row, user_data):
    """키워드_리스트 + 주요 제품 + 사업목적 + 10차산업코드명을 ';'로 묶어 임베딩 입력 텍스트 생성.

    주요 제품/사업목적은 normalize_user_input 으로 명사 키워드만 추출해 추가한다.
    10차산업코드명은 기존 키워드_리스트 길이가 3 이하일 때만, 역시 normalize_user_input
    으로 명사 키워드만 추출해(문장 통째 아님) 추가한다.
    """
    user_data = user_data or {}
    products = (user_data.get('main_products', '') or '').strip()
    purpose = (user_data.get('business_purpose', '') or '').strip()
    return _join_combined(row, [products, purpose], fallback_col='10차산업코드명')


def build_combined_project_text(row, user_data):
    """키워드_리스트 + 한글 키워드 + 과학기술표준분류 + 과제명을 ';'로 묶어 임베딩 입력 텍스트 생성.

    한글 키워드/과학기술표준분류/과제명 모두 normalize_user_input 으로 명사 키워드만
    추출해 추가한다(문장 통째 아님). 과제명은 키워드_리스트 길이와 무관하게 **항상**
    포함한다(기존 키워드와 중복되는 토큰은 한 번만 사용). 순서: 한글키워드 → 과기표준분류 → 과제명.
    """
    user_data = user_data or {}
    kr_kw = (user_data.get('korean_keywords', '') or '').strip()
    sci_class = (user_data.get('sci_tech_class', '') or '').strip()
    title_raw = row.get('과제명') if hasattr(row, 'get') else None
    try:
        title = '' if title_raw is None or pd.isna(title_raw) else str(title_raw).strip()
    except (TypeError, ValueError):
        title = str(title_raw).strip()
    return _join_combined(row, [kr_kw, sci_class, title], fallback_col=None)
