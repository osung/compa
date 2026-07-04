import gc
import json
import os
import re
import time
from pathlib import Path

import pandas as pd
import torch
from transformers import AutoTokenizer
from vllm import LLM, SamplingParams

'''프롬프트 구성'''

import ast
import json
import numpy as np
import pandas as pd

def to_py(x):
    if isinstance(x, dict):
        return {k: to_py(v) for k, v in x.items()}
    if isinstance(x, list):
        return [to_py(v) for v in x]
    if isinstance(x, np.generic):
        return x.item()
    return x

def clean_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    if s.lower() == "none":
        return ""
    return s

def parse_struct(x):

    if x is None:
        return None

    # pandas NaN 처리
    if isinstance(x, float) and pd.isna(x):
        return None

    if isinstance(x, (list, dict)):
        return x

    if isinstance(x, str):
        x = x.strip()
        if not x:
            return None

        # JSON 문자열 시도
        try:
            return json.loads(x)
        except Exception:
            pass

        # Python literal 문자열 시도
        try:
            return ast.literal_eval(x)
        except Exception:
            return x

    return x


def extract_group(x):
    x = parse_struct(x)

    if not isinstance(x, dict):
        return None

    group = x.get("group")

    if group is None:
        return None

    s = str(group).strip()
    if not s or s.lower() in {"nan", "none", "null"}:
        return None

    return s


def normalize_conduct_list_company(x):
    """
    최종 목표:
    [
        {
            "company_id": "...",
            "company_name": "...",
            "company_info": {...}
        },
        ...
    ]
    """
    x = parse_struct(x)

    if not isinstance(x, list):
        return []

    result = []

    for item in x:
        if not isinstance(item, dict):
            continue

        info = parse_struct(item.get("company_info"))
        if not isinstance(info, dict):
            info = {}

        company_id = clean_text(item.get("company_id"))
        fallback_name = clean_text(item.get("company_name"))

        company_info = {
            "company_name": clean_text(fallback_name),
            "region": clean_text(info.get("region", None)),
            "company_keyword": normalize_str_list(info.get("company_keyword", [])),
            "company_purpose_list": normalize_str_list(info.get("company_purpose_list", [])),
            "company_patent": normalize_str_list(info.get("company_patent", [])),
        }

        has_info = any([
            company_info["company_name"],
            company_info["region"],
            company_info["company_keyword"],
            company_info["company_purpose_list"],
            company_info["company_patent"],
        ])

        if company_id or has_info:
            result.append({
                "company_id": company_id,
                "company_name": company_info["company_name"],
                "company_info": company_info
            })

    return result


def normalize_conduct_list_project(x):
    """
    최종 목표:
    [
        {
            "project_id": "...",
            "project_name": "...",
            "project_info": {...}
        },
        ...
    ]
    """
    x = parse_struct(x)

    if not isinstance(x, list):
        return []

    result = []

    for item in x:
        if not isinstance(item, dict):
            continue

        info = parse_struct(item.get("project_info"))
        if not isinstance(info, dict):
            info = {}

        project_id = clean_text(item.get("project_id"))
        fallback_name = clean_text(item.get("project_name") or item.get("과제명"))

        project_info = {
            "project_name": clean_text(info.get("project_name") or info.get("과제명") or fallback_name),
            "project_keyword": normalize_str_list(info.get("project_keyword") or info.get("키워드_project") or []),
            "paper": normalize_str_list(info.get("paper", None)),
            "patent": normalize_str_list(info.get("patent", None)),
        }

        has_info = any([
            project_info["project_name"],
            project_info["project_keyword"],
            project_info["paper"],
            project_info["patent"],
        ])

        if project_id or has_info:
            result.append({
                "project_id": project_id,
                "project_name": project_info["project_name"],
                "project_info": project_info
            })

    return result


def normalize_str_list(x):
    """
    문자열/리스트/문자열화된 리스트를
    ['a', 'b', ...] 형태로 정규화
    """
    x = parse_struct(x)

    if x is None:
        return []

    if not isinstance(x, list):
        x = [x]

    result = []
    for v in x:
        if v is None:
            continue
        s = str(v).strip()
        if s:
            result.append(s)

    return result

REGION_GROUPS = {
    "수도권": ["서울", "경기", "인천"],
    "충청권": ["충남", "충북", "대전", "세종"],
    "호남권": ["전북", "전남", "광주"],
    "영남권": ["경남", "경북", "대구", "부산"],
    "강원권": ["강원"],
}

REGION_TO_GROUP = {}
for group_name, regions in REGION_GROUPS.items():
    for r in regions:
        REGION_TO_GROUP[r] = group_name


def normalize_region(x):
    if x is None:
        return ""
    return str(x).strip()


def analyze_region_relation(company_region, conduct_list_company):
    company_region = normalize_region(company_region)
    company_group = REGION_TO_GROUP.get(company_region, "")

    result = {
        "company_region": company_region,
        "company_region_group": company_group,
        "same_region_count": 0,
        "similar_region_count": 0,
        "same_region_companies": [],
        "similar_region_companies": [],
        "same_region_names": [],
        "similar_region_names": [],
        "same_region_group": "",
        "similar_region_group": "",
        "has_same_region": False,
        "has_similar_region": False,
    }

    if not company_region:
        return result

    same_region_companies = []
    similar_region_companies = []

    for item in conduct_list_company:
        if not isinstance(item, dict):
            continue

        info = item.get("company_info", {})
        if not isinstance(info, dict):
            info = {}

        c_name = clean_text(item.get("company_name") or info.get("company_name"))
        c_region = normalize_region(info.get("region"))

        if not c_region:
            continue

        company_item = {
            "company_name": c_name,
            "region": c_region
        }

        # 1) 정확히 같은 지역
        if c_region == company_region:
            same_region_companies.append(company_item)
            continue

        # 2) 강원은 exact only
        if company_region == "강원" or c_region == "강원":
            continue

        # 3) 같은 권역이면 유사 지역
        c_group = REGION_TO_GROUP.get(c_region, "")
        if company_group and c_group and company_group == c_group:
            similar_region_companies.append(company_item)

    result["same_region_count"] = len(same_region_companies)
    result["similar_region_count"] = len(similar_region_companies)
    result["same_region_companies"] = same_region_companies
    result["similar_region_companies"] = similar_region_companies
    result["same_region_names"] = [x["company_name"] for x in same_region_companies if x["company_name"]]
    result["similar_region_names"] = [x["company_name"] for x in similar_region_companies if x["company_name"]]
    result["same_region_group"] = company_group if same_region_companies else ""
    result["similar_region_group"] = company_group if similar_region_companies else ""
    result["has_same_region"] = len(same_region_companies) > 0
    result["has_similar_region"] = len(similar_region_companies) > 0

    return result


SYSTEM_PROMPT = (
    "너는 추천 시스템의 '추천 이유'를 설명하는 한국어 AI야. "
    "입력 JSON의 정보를 기반으로 특정 회사에 특정 과제가 왜 추천되었는지 한국어로 설명한다. "
    "기술적 원리나 설명이 필요한 경우에는 일반적인 산업/기술 지식을 활용해 쉽게 풀어 설명할 수 있다. "
    "다만 입력 정보와 무관한 새로운 사실(특정 기업의 추가 사업, 특정 수치, 특정 기술 보유 등)을 만들어서는 안 된다."
    "아래 규칙을 기반으로 최종 결과만 생성해. 중간 추론 과정은 출력하지 마.\n"
    "1) 회사 정보 해석\n"
    "   1-1) company.purpose는 유효한 값이 있는 경우에만 해석하되, company.keyword와 기술적으로 연결되는 항목만 선택적으로 사용하며, 연결성이 낮은 항목은 설명에서 제외한다.\n"
    "   1-2) company.keyword는 유효한 값이 있는 경우에만 의미 단위로 묶어 회사의 핵심 역량과 기술 요소를 도출한다.\n"
    "   1-3) company.company_patent는 유효한 값이 있는 경우에만 검토하여 회사의 기술 축, 구현 방식, 적용 가능 분야를 파악한다. 연관성 서술에 활용할 때는 회사가 '보유한 특허'임을 분명히 드러내되(예: '~ 관련 특허를 보유'), 특허 제목 전체는 그대로 나열하지 않고 기술 주제로 묶어 표현한다.\n"
    "2) 과제 정보 해석\n"
    "   2-1) project.title, project.keywords는 유효한 값이 있는 경우에만 과제의 핵심 대상과 기술 방향을 추출한다.\n"
    "   2-2) related_research.paper 또는 related_research.patent에 유효한 값이 있는 경우에만 집중하는 핵심 기술 또는 연구 방향을 파악한다.\n"
    "3) 1)과 2)의 요약 내용을 근거로 과제와 회사의 연관성을 '연관성' 섹션에서 설명한다.\n"
    "   - 과제 추천 방향이므로, 먼저 추천 과제의 핵심 목적·대상·기술 방향을 설명한 뒤, 그 과제의 기술이 회사의 사업·기술·제품·제조 공정·연구개발과 어떻게 연결되는지 이어서 설명한다(과제 설명이 먼저, 기업 설명이 나중).\n"
    "   - 연결성 설명에는 반드시 1문장 이상으로 과제의 핵심 기술/방법이 왜 필요한지(작동 원리, 메커니즘, 해결하려는 문제의 원인)를 일반적인 기술 지식으로 풀어서 설명한다.\n"
    "   - 기술 설명은 '조건 또는 변수 → 그 변화로 발생하는 문제 또는 결과 → 그래서 해당 기술이 필요함'의 인과 구조로 설명한다.\n"
    "   - 연결성 설명은 다음 순서를 따른다: (추천 과제의 핵심 목적·기술과 그것이 필요한 이유) → (해당 기술이 실제로 활용되거나 적용되는 중간 단계) → (회사의 기술/사업과의 연결) \n"
    "   - 회사명은 반드시 입력 JSON의 company.name 값을 그대로 사용한다.\n"
    "   - related_research.paper(과제의 논문 성과) 또는 related_research.patent(과제의 특허 성과)에 유효한 값이 있으면, 그 성과가 드러내는 과제의 핵심 기술·연구 방향을 회사의 기술·사업 역량과 연결지어 연관성 근거로 함께 서술한다. 다만 논문·특허 제목 전체를 그대로 나열하지 말고 공통 기술 주제로 묶어 표현하며, 값이 없으면 언급하지 않는다.\n"
    "4) 기술 설명이 필요한 경우에는 일반적인 산업 공정 지식이나 기술 지식을 활용하여 설명할 수 있다.\n"
    "   다만 입력 JSON에 없는 회사 사실이나 과제 정보를 새로 만들어서는 안 된다.\n"
    "5) 기업의 재무 및 기술 역량을 바탕으로 '추천 기업 우수성' 섹션에서 작성한다.\n"
    "    - 이 섹션은 '추천된 기업 자체의 우수성'만 다룬다. 과제의 총연구비·논문·특허 등 '과제 우수성'은 이 섹션에 어떠한 형태로도 작성하지 않는다(과제 정보는 '연관성' 섹션에서만 활용).\n"
    "    - company.벤처기업여부, company.이노비즈여부, company.메인비즈여부, company.ASTI 여부, company.특구 여부 중 값이 'Y'인 항목만 선택하여 회사의 인증/지정 근거로 반영한다.\n"
    "    - company.patent_count 값이 있고 0보다 크면 '총 ○건의 특허를 보유' 처럼 기업의 기술 역량·혁신성 근거로 반영한다. company.patent_matching_related 에 특허 제목이 있으면 대표 1~3개가 공통적으로 가리키는 기술 분야로 묶어 어떤 기술의 특허를 보유했는지 간략히 덧붙인다(제목을 길게 그대로 나열하지 말 것). patent_count 가 없거나 0이면 특허 보유 관련 내용을 작성하지 않는다.\n"
    "    - company.매출성장율_상위비율과 company.매출성장율_group 값이 모두 있을 때만 언급한다.\n"
    "    - company.매출성장율_group이 '전체'이면 '매출성장율이 전체의 상위 n% 수준'으로, 그 외에는 '매출성장율이 {group} 업종 내 상위 n% 수준'으로 표현한다.\n"
    "    - company.영업이익율_상위비율과 company.영업이익율_group 값이 모두 있을 때만 언급한다.\n"
    "    - company.영업이익율_group이 '전체'이면 '영업이익율이 전체의 상위 n% 수준'으로, 그 외에는 '영업이익율이 {group} 업종 내 상위 n% 수준'으로 표현한다.\n"
    "    - company.연구개발비_상위비율과 company.연구개발비_group 값이 모두 있을 때만 언급한다.\n"
    "    - company.연구개발비_group이 '전체'이면 '연구개발비가 전체의 상위 n% 수준'으로, 그 외에는 '연구개발비가 {group} 업종 내 상위 n% 수준'으로 표현한다.\n"
    "    - company.부채비율_하위비율과 company.부채비율_group 값이 모두 있을 때만 언급한다.\n"
    "    - company.부채비율_group이 '전체'이면 '부채비율이 전체의 상위 n% 수준'으로, 그 외에는 '부채비율이 {group} 업종 내 상위 n% 수준'으로 표현한다.\n"
    "    - 위 비율 정보와 group은 해당 값이 모두 있을 때만 언급하고, 값이 없으면 완전히 생략한다.\n"
    "    - 정량 지표는 숫자를 나열하듯 쓰지 말고, 기업의 우수성을 설명하는 보조 근거로 자연스럽게 종합 서술한다.\n"
    "    - 추천 기업 우수성 섹션은 다음 순서로 작성한다.\n"
    "      1. company.* 인증/지정 항목 중 값이 'Y'인 항목이 있는 경우 회사의 기술·사업 역량 근거로 설명한다.\n"
    "      2. company.patent_count 가 0보다 크면 보유 특허 실적(총 건수 + 대표 기술 분야)을 기업의 기술 역량 근거로 서술한다.\n"
    "      3. company.* 재무지표가 존재하는 경우 회사의 성장성·수익성·재무 안정성 근거로 자연스럽게 종합 서술한다.\n"
    "      4. 마지막으로 기업 자체의 역량 및 추천 타당성을 종합 결론으로 작성한다(과제 우수성은 제외).\n"
    "6) 다음 내용을 바탕으로 '유사 사례' 섹션을 생성한다.\n"
    "   - '유사 사례'는 두 축으로 구성한다: (가) 기업의 보유특허·수행 과제와 추천 과제의 논문·특허 성과 사이의 기술적 연관성, (나) 추천(매칭)된 과제를 수행한 기업들과 현재 기업의 유사성. 특히 (나)는 '현재 기업과 유사한 기업들이 이미 이 과제를 수행했다'는 협업 필터링 관점의 추천 근거로 서술한다.\n"
    "   - (단락 구성·필수) '유사 사례 및 실적' 섹션은 반드시 두 단락으로 나누어 작성하고, 두 단락 사이는 빈 줄(줄바꿈 문자 두 개 '\\n\\n')로 구분한다. 첫째 단락에는 (가) 추천 과제의 논문·특허 성과와 기업 보유특허의 기술적 연관성을, 둘째 단락에는 (나) 추천 과제를 수행한 기업(과제수행기관)과 현재 기업의 유사성(및 지역적 연관성)을 작성한다. (가) 또는 (나) 중 한쪽 내용이 없으면 해당 단락은 생략하고, 작성되는 단락이 하나뿐이면 단락 구분 줄바꿈도 넣지 않는다.\n"
    "  [매칭 과제와의 유사성]\n"
    "   - company.has_patent_matching_related가 true인 경우에만, patent_matching_related에 포함된 특허 제목들과 2)에서 파악한 현재 추천 과제의 내용을 비교하여 기술적 연관성을 설명한다.\n"
    "   - 이때 해당 내용이 회사가 '보유한 특허'임을 문장에서 분명히 드러낸다(예: '~ 관련 특허를 보유하여'). 다만 특허 제목 전체를 그대로 나열하지는 말고, 특허들이 공통적으로 가리키는 기술 주제, 적용 분야, 해결하려는 문제와 현재 추천 과제의 목표·핵심 기술·적용 방식이 어떻게 연결되는지 종합하여 설명한다.\n"
    "   - company.has_patent_matching_related가 false이면 특허 매칭 관련 내용을 어떠한 형태로도 작성하지 않는다.\n"

    " [매칭 과제의 유사 논문 및 특허 성과]\n"
    "    - has_paper 또는 has_patent 중 하나 이상이 true인 경우에만 매칭 과제의 유사 논문 및 특허 성과 내용을 작성한다.\n"
    "    - has_paper와 has_patent가 모두 false이면 매칭 과제의 유사 논문 및 특허 성과 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "    - has_paper가 false이면 논문 실적, 논문 부재, 논문 기반 기술 연관성 관련 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "    - has_patent가 false이면 특허 실적, 특허 부재, 특허 기반 기술 연관성 관련 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "    - 논문 또는 특허가 여러 개 제공되는 경우 각 항목을 단순 나열하지 말고 반복되는 기술 주제 또는 공통 연구 방향을 중심으로 종합적으로 설명한다.\n"
    "    - has_paper 또는 has_patent 중 true인 항목의 내용만 종합하여 해당 과제가 어떤 기술 분야나 연구 방향에 집중하고 있는지 정리하고 회사 목적 및 기술과 어떻게 연관 되는지 설명한다.\n"
    "    - 이후 해당 기술이 일반적으로 어떤 장치나 시스템에서 사용되는지 설명하고, 그 장치 또는 시스템이 1)에서 파악한 현재 회사의 사업과 어떻게 연결되는지 설명한다.\n"
    "    - has_paper 또는 has_patent 중 하나 이상이 true인 경우, 과제의 기술 → 기술이 필요한 이유 → 기술이 사용되는 장치 또는 시스템 → 회사 사업과의 연관 근거 순서로 설명한다.\n"

    "   - conduct_list_company와 conduct_list_project는 각각 독립적으로 판단한다.\n"
    "   - 빈 리스트, None, 누락된 값에 대해서는 어떤 형태로도 언급하지 않는다.\n"
    "   - 빈 값을 근거로 한 추측, 보완 설명, 일반화된 설명을 생성하지 않으며, '없다', '제공되지 않았다', '비어 있다', '확인되지 않는다'와 같은 표현은 절대 생성하지 않는다.\n"
    "   - (중요) 수행 기업·지역 연관성 등 어떤 정보가 없을 때, 그 부재나 생략 사실 자체를 문장으로 서술하지 않는다. '정보가 제공되지 않아', '정보가 없어', '해당 내용은 생략한다', '생략합니다', '활용되지 않는다/않으나', '작성할 수 없다', '수행 기업이 존재하지 않아' 같은 표현이나 conduct_list_company 같은 내부 필드명은 절대 출력하지 않는다. 근거가 없는 부분은 그냥 작성하지 말고, 근거가 있는 내용만 자연스럽게 서술한 뒤 문장을 끝낸다.\n"

    "   [추천된 과제를 수행한 기업 유사 사례]\n"
    "   - has_conduct_company가 True인 경우에만 추천된 과제를 수행한 기업 유사 사례 내용을 작성한다.\n"
    "   - has_conduct_company가 False이면 수행 기업, 수행 기업 부재, 수행 기업 유사 사례 관련 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "   - (중요·먼저) 이 블록의 회사들은 '추천(매칭)된 과제를 실제로 수행한 기업'이다. 이 블록은 둘째 단락의 맨 앞에 오며, 반드시 그 첫 문장을 '과제수행기관인 {company_name}은 …' 형태로 시작해 그 회사가 추천 과제를 수행한 기관임을 먼저 명시한 뒤 비교 설명을 이어간다. 회사 이름만 단독으로(예: '{company_name}은 …') 시작하거나, '이 과제와 유사한 기술을 보유한 기업', '유사한 사업을 수행하는 기업' 처럼 수행 사실을 흐리는 표현은 쓰지 않는다.\n"
    "   - conduct_list_company에서는 각 항목의 company_info 안에 있는 company_name, company_purpose_list, company_keyword, region, company_patent(수행기관 보유특허), conduct_projects(수행기관이 수행한 과제) 중 값이 존재하는 정보만 활용한다.\n"
    "   - 특히 수행기관의 company_patent·conduct_projects 와 현재(매칭) 기업의 보유특허(patent_matching_related)·수행과제(conduct_list_project)를 비교하여, 두 기업이 어떤 기술 주제의 특허·수행과제에서 공통되는지 분석한다(협업 필터링: 매칭 기업과 유사한 특허·수행과제를 가진 기업이 이미 이 과제를 수행함을 근거로 제시).\n"
    "   - 추천된 과제를 수행한 기업(수행 기업)의 company_info(보유특허·수행과제·키워드·사업목적·지역 등)와 1)에서 파악한 현재(분석/추천) 기업의 정보를 '두 기업끼리만' 비교하여, 두 기업이 어떤 기술·사업·공정·제품·연구개발 방향에서 유사한지 설명한다.\n"
    "   - (금지) 수행 기업과 '추천(매칭)된 과제' 사이의 연관성(수행 기업이 그 과제를 수행했다는 사실, 수행 기업 기술이 과제와 맞다는 식의 설명)은 자명하므로 절대 분석하거나 언급하지 않는다. 이 블록의 비교 대상은 오직 '수행 기업 ↔ 현재(분석/추천) 기업'이며, 추천 과제와의 연관성은 다른 블록에서만 다룬다.\n"
    "   - 단순히 '유사하다'고 표현하지 말고, 수행 기업의 기술 또는 사업 내용이 현재(분석/추천) 기업과 어떤 점에서 공통적으로 닮아 있는지 구체적으로 서술한다.\n"
    "   - 공통점이 분명한 경우, '현재 기업과 유사한 기업들이 이미 이 과제를 수행했다'는 점을 추천의 핵심 근거(협업 필터링 관점)로 분명히 제시한다.\n"
    "   - 단, 수행 회사의 기술·사업·공정·제품·연구개발 방향이 현재 회사 및 추천 과제와 실질적인 공통점이 뚜렷하지 않으면(연관성이 낮으면) 억지로 유사하다고 엮지 말고, 해당 수행 회사에 대한 내용을 유사 사례에 포함하지 않는다(언급 자체를 생략한다). 공통점이 분명한 수행 회사가 하나도 없으면 '추천된 과제를 수행한 기업 유사 사례' 자체를 작성하지 않는다.\n"
    "   - conduct_list_company에 여러 항목이 있을 경우 각 회사를 하나씩 단순 나열하지 말고, 반복되는 공통 기술 주제, 공통 사업 방향, 공통 문제 해결 방식 중심으로 종합하여 설명한다.\n"

    "   [추천 과제를 수행한 기업의 지역적 연관성]\n"
    "   - 지역 근거는 has_conduct_company가 True인 경우에만 보조 근거로 활용할 수 있다.\n"
    "   - has_conduct_company가 False이면 지역적 연관성 관련 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "   - 지역 근거는 기술적 유사성의 대체가 아니라 보조 근거로만 사용하며, 같은 지역 또는 유사 지역이라는 이유만으로 기술적 적합성을 단정하지 않는다.\n"
    "   - conduct_list_company의 각 항목에 포함된 company_info.region과 company.region을 비교하여 지역적 연관성을 판단한다.\n"
    "   - 현재 회사와 동일한 지역의 수행 기업이 있으면 유사 지역보다 우선적으로 설명한다.\n"
    "   - 동일 지역 기업이 여러 개이면 company.region 값을 직접 언급하여 지역적 연관성이 비교적 강한 보조 근거임을 설명한다.\n"
    "   - 동일 지역 기업이 없고 유사 지역 기업이 있을 때만 유사 지역 근거를 사용한다.\n"
    "   - 유사 지역은 수도권(서울, 경기, 인천), 충청권(충남, 충북, 대전, 세종), 호남권(전북, 전남, 광주), 영남권(경남, 경북, 대구, 부산), 강원권(강원) 기준으로 판단한다.\n"
    "   - 강원권은 정확히 '강원'이 일치할 때만 지역 근거로 사용한다.\n"
    "   - region_relation.has_same_region이 true이면 동일 지역 근거를 보조적으로 사용하고, company.region 값을 직접 언급한다.\n"
    "   - region_relation.has_same_region이 false이고 region_relation.has_similar_region이 true이면 유사 권역 근거를 보조적으로 사용하고, region_relation.company_region_group 값을 직접 언급한다.\n"

    "   [기업이 수행한 과제 유사 사례]\n"
    "   - has_conduct_project가 True인 경우에만 기업이 수행한 과제 유사 사례 내용을 작성한다.\n"
    "   - has_conduct_project가 False이면 수행 과제, 수행 이력, 과거 과제 부재와 관련된 내용을 어떠한 형태로도 작성하지 않는다.\n"
    "   - conduct_list_project에서는 각 항목의 project_info 안에 있는 project_name, project_keyword, paper, patent 중 값이 존재하는 정보만 활용한다.\n"
    "   - conduct_list_project에 여러 항목이 있을 경우 각 과제를 하나씩 단순 나열하지 말고, 공통 기술 주제, 공통 연구개발 방향, 공통 문제 해결 방식 중심으로 종합하여 설명한다.\n"
    "   - 현재 회사가 과거에 수행했던 과제들의 project_info와 2)에서 파악한 현재 추천된 과제의 내용을 비교하여 목표, 핵심 기술, 적용 방식, 해결하려는 문제 측면에서 어떤 유사성이 있는지 설명한다.\n"

    "7) 모든 요소를 섹션 구조로 일관되게 정리한다.\n"
    "8) 모든 섹션은 일반인이 이해할 수 있는 수준으로 작성한다.\n\n"
    "[출력 규칙]\n"
    "- 위 내부 사고 단계나 중간 판단, 계산 과정은 절대 출력하지 않는다.\n"
    "- 추측하거나 없는 사실을 만들지 않는다.\n"
    "- 입력 JSON에 값이 없는 항목은 언급하지 않는다.\n"
    "- 다만 기술 설명이나 원리 설명이 필요한 경우에는 일반적인 산업 또는 기술 지식을 활용하여 설명할 수 있다.\n"
    "- 모든 설명은 '평가'가 아니라 '추천 이유 설명'의 관점에서 작성한다.\n"
    )


FEWSHOT_MESSAGES = [
    {"role": "user", "content": (
        "아래 JSON만을 근거로 추천 근거를 작성해.\n"
        "출력은 JSON 하나만 반환해. (JSON 외 텍스트 금지)\n"
        "키는 output_requirements.format에 있는 섹션 제목을 그대로 사용해.\n\n"
        + json.dumps({
            "company": {
                "company_id": "C001",
                "name": "샘플회사_알파12a",
                "purpose": ['반도체 및 평판디스플레이 제조용 기계 제조업', '항공기, 우주선 및 보조장치 제조업', '공학 연구개발업'],
                "keyword": [
                        '코팅공정용', '나노물질자가정렬방법', '코팅물질', '용액공정', '용액기반', '공압모듈', '코팅장비마스크패터닝', '코팅장비', '코팅솔루션', '필름제조용', '금속입자소결체', '진공증착', '공압부품', '잉크토출장치', '공정기술', '피인쇄물질', '밸브제작', '용액'
                    ],
                "patent_matching_related": [
                        {"특허명":"프린팅 장치"},
                        {"특허명":"유도보조 전극을 포함하는 유도 전기수력학 젯 프린팅 장치"},
                        {"특허명":"피드백 제어형 인쇄 시스템"},
                        {"특허명":"전기수력학 방식의 분사 노즐"}
                ],
                "has_patent_matching_related": True,
                "patent_count": 24,

                "conduct_list_project": [
                    {
                        "과제명": "고분자 기반 기능성 필름 제조 공정 개발",
                        "과제코드": "P9001",
                        "project_info": {
                            "project_name": "고분자 기반 기능성 필름 제조 공정 개발",
                            "project_keyword": ["고분자", "필름", "코팅", "건조", "표면 제어"],
                            "paper": [],
                            "patent": []
                        }
                    },
                    {
                        "과제명": "정밀 프린팅 기반 소재 패터닝 기술 개발",
                        "과제코드": "P9002",
                        "project_info": {
                            "project_name": "정밀 프린팅 기반 소재 패터닝 기술 개발",
                            "project_keyword": [],
                            "paper": None,
                            "patent": None
                        }
                    }
                ],
                "has_conduct_project": True,

                "region": "서울",

                "벤처기업여부": "Y",
                "이노비즈여부": "Y",
                "메인비즈여부": "N",
                "ASTI 여부": "N",
                "특구 여부": "N",
                "매출성장율_상위비율": '5',
                "매출성장율_group": "전체",

                "영업이익율_상위비율": "",
                "영업이익율_group": None,

                "부채비율_하위비율": '10',
                "부채비율_group": "전체",

                "연구개발비_상위비율": '20',
                "연구개발비_group": "전문직별 공사업",

            },
            "project": {
                "project_id": "P001",
                "title": "액정 엘라스토머 기반 4D 프린팅 소재 개발",
                "keyword": [
                        "코팅공정용", "용액공정", "금속입자소결체", "잉크토출장치", "공정기술"
                    ],
                "project_score": 70.21047,
                "cosine_distance": 0.024936,
                "description": "이 과제는 테스트기관베타7X에서 수행했으며, 주조/용접/접합 분야에 속합니다. 또한, 총연구비는 전체의 상위 5%에 속하며, 총연구기간은 8개월 30일입니다. 이 과제는 액정 엘라스토머 소재를 기반으로 시간이 흐르면서 형태가 변하는 4D 프린팅 소재를 개발하는 연구입니다.이 소재는 고분자 탄성 액추에이터와 관련된 4D 프린팅 기술을 활용하며, 액정 엘라스토머 필름 제조 방법 등 기능성 소재의 제작 기술을 다룹니다.논문과 특허를 보면 형태 변화가 가능한 소재와 그 제조 기술에 대한 연구가 중심입니다.",
                "총연구비_상위비율": '5',
                "총연구비_group": "전체"
            },

            "conduct_list_company": [
                {
                    "company_name": "가상회사_유동제어_1",
                    "company_id": "C101",
                    "company_info": {
                        "company_name": "가상회사_유동제어_1",
                        "region": "서울",
                        "company_keyword": ["기능성 필름", "정밀 코팅", "박막 형성", "공정 제어"],
                        "company_purpose_list": ["디스플레이 및 전자재료용 소재 개발"],
                        "company_patent": []
                    }
                },
                {
                    "company_name": "임시조직a12k",
                    "company_id": "C102",
                    "company_info": {
                        "company_name": "임시조직a12k",
                        "region": "경기",
                        "company_keyword": [],
                        "company_purpose_list": [],
                        "company_patent": []
                    }
                }
            ],
            "has_conduct_company": True,

            "region_relation": {
              "company_region": "서울",
              "company_region_group": "수도권",
              "same_region_count": 1,
              "similar_region_count": 1,
              "same_region_companies": [
                  {"company_name": "가상회사_유동제어_1", "region": "서울"}
              ],
              "similar_region_companies": [
                  {"company_name": "임시조직a12k", "region": "경기"}
              ],
              "same_region_names": ["가상회사_유동제어_1"],
              "similar_region_names": ["임시조직a12k"],
              "same_region_group": "수도권",
              "similar_region_group": "수도권",
              "has_same_region": True,
              "has_similar_region": True
          },
            "related_research": {
                "paper": ["습윤 고분자 탄성 액추에이터의 4D 프린팅"],
                "patent": ["폴리로탁산 가교체를 도입한 액정 엘라스토머 필름의 제조 방법"]
            },
            "paper_list_count" : 1,
            "patent_list_count" : 1,
            "has_paper" : True,
            "has_patent" : True,
            "output_requirements": {
                "language": "ko",
                "format": [
                    "연관성",
                    "추천 기업 우수성",
                    "유사 사례 및 실적"
                ],
                "section_sentence_range": "각 섹션 4~6문장",
                "forbidden": ["예시", "참고", "제출", "메타"]
            }
        }, ensure_ascii=False)
    )},

    {"role": "assistant", "content": json.dumps({
        "연관성": (
            "샘플회사_알파12a는 코팅공정, 용액기반 공정, 잉크 토출 장치, 공압 모듈 및 정밀 코팅 장비와 관련된 기술 요소를 보유하고 있어 액정 엘라스토머 기반 4D 프린팅 소재 개발 과제와 높은 기술적 연결성을 가진다."
            "액정 엘라스토머 기반 4D 프린팅은 소재 내부 배열과 코팅 균일도가 변하면 형태 변화 성능과 구동 안정성이 달라질 수 있기 때문에, 용액 공정과 정밀 코팅 기술을 통해 소재의 두께와 배열을 안정적으로 제어하는 기술이 중요하다."
            "샘플회사_알파12a의 코팅장비, 공정기술, 잉크토출장치 관련 역량은 기능성 필름 제조와 정밀 소재 패터닝 과정에서 활용될 수 있으며, 이는 액정 엘라스토머 필름 제조 및 4D 프린팅 공정의 중간 제조 단계와 직접 연결된다."
            "특히 금속입자소결체와 용액기반 공정 기술은 기능성 소재를 균일하게 형성하고 패턴 정밀도를 높이는 데 활용될 수 있어 고분자 기반 액추에이터 제작 공정과의 연관성이 높다."
            "또한 프린팅 장치, 피드백 제어형 인쇄 시스템, 전기수력학 방식의 분사 노즐과 같은 특허들은 정밀 분사 및 인쇄 공정 안정화와 관련된 기술 축을 형성하고 있으며, 이는 4D 프린팅 소재 제조 과정에서 필요한 정밀 패터닝 및 균일 코팅 기술과 자연스럽게 연결된다."
        ),
        "추천 기업 우수성": (
            "샘플회사_알파12a는 벤처기업과 이노비즈 인증을 보유하고 있어 기술 기반 사업화 역량을 갖춘 기업으로 볼 수 있습니다."
            "또한 총 24건의 특허를 보유하고 있으며, 정밀 분사·전기수력학 프린팅 등 토출·인쇄 공정 제어 분야의 특허를 중심으로 기술 역량을 확보하고 있습니다."
            "재무적으로도 매출성장율이 전체의 상위 5% 수준이고, 부채비율이 전체의 상위 10% 수준이며, 연구개발비가 전문직별 공사업 업종 내 상위 20% 수준으로 나타나 성장성·재무 안정성·연구개발 투입 측면에서 우수합니다."
            "종합하면 이 기업은 인증과 특허 기반의 기술 역량에 견조한 재무 지표를 함께 갖추고 있어 추천 타당성이 높은 기업입니다.",
        ),
        "유사 사례 및 실적": (
            "추천된 과제는 본 과제 수행을 통해 습윤 고분자 탄성 액추에이터·액정 엘라스토머 필름 제조 관련 논문·특허를 확보했으며, 이는 형태 변화형 고분자 소재와 기능성 필름 제작 분야의 성과입니다."
            "샘플회사_알파12a가 보유한 프린팅 장치·전기수력학 분사 노즐 등 정밀 토출·패턴 형성 관련 특허는 이러한 고분자 소재를 균일하게 제어·분사하는 기술과 맞닿아 있어, 과제의 연구성과와 기술적으로 연관됩니다.\n\n"
            "과제수행기관인 가상회사_유동제어_1은 기능성 필름·정밀 코팅·박막 형성·공정 제어 기술을 보유하고 있으며, 이는 현재 기업인 샘플회사_알파12a의 코팅·증착·소재 제조 역량과 직접적으로 공통됩니다."
            "두 기업 모두 박막·코팅 공정과 소재 제어 기술을 사업 기반으로 삼고 있어, 현재 기업과 유사한 기업이 이미 이 과제를 수행했다는 점(협업 필터링)이 추천을 뒷받침합니다."
            "또한 과제수행기관 가상회사_유동제어_1은 현재 기업과 동일한 서울 지역에 위치해 지역적 연관성도 보조 근거가 됩니다."
        ),
    }, ensure_ascii=False)}
]



def build_company_single_prompt(company_row, rec_row):
    c_id = company_row["company_id"]
    c_name = company_row["company_name"]
    #c_desc = company_row.get("company_description", "")
    c_purpose = normalize_str_list(company_row.get("company_purpose_list", []))
    c_keyword = normalize_str_list(company_row.get("키워드_company", []))

    pid = rec_row["project_id"]
    pname = rec_row["project_name"]
    pscore = rec_row.get("project_score", "")
    dist = rec_row.get("cosine_distance", "")
    #p_desc = rec_row.get("project_description", "")
    keyword_proj = normalize_str_list(rec_row.get("키워드_project", []))

    paper_list = normalize_str_list(rec_row.get("paper", []))
    patent_list = normalize_str_list(rec_row.get("patent", []))
    paper_list_count = len(paper_list)
    patent_list_count = len(patent_list)
    has_paper = len(paper_list) > 0
    has_patent = len(patent_list) > 0

    # 수행 과제/기업 파싱
    conduct_list_company = normalize_conduct_list_company(
        rec_row.get("conduct_list_company", [])
    )

    conduct_list_project = normalize_conduct_list_project(
        company_row.get("conduct_list_project", [])
    )
    has_conduct_company = len(conduct_list_company) > 0
    has_conduct_project = len(conduct_list_project) > 0

    company_region = str(company_row.get("region", "")).strip()
    region_relation = analyze_region_relation(company_region, conduct_list_company)

    # company_patent_sim 파싱
    # patent_matching_related 파싱
    patent_matching_related = company_row.get("company_patent_sim", [])
    if isinstance(patent_matching_related, str):
        try:
            patent_matching_related = ast.literal_eval(patent_matching_related)
        except:
            patent_matching_related = []

    if not isinstance(patent_matching_related, list):
        patent_matching_related = []

    valid_patent_matching_related = []
    for x in patent_matching_related:
        # dict 형태
        if isinstance(x, dict):
            title = x.get("특허명", "")
            if str(title).strip():
                valid_patent_matching_related.append({
                    "특허명": str(title).strip()
                })

        # 문자열 형태
        elif isinstance(x, str):
            title = x.strip()
            if title:
                valid_patent_matching_related.append({
                    "특허명": title
                })


    # fewshot과 같은 format 규칙
    fmt = [
        "연관성",
        "추천 기업 우수성",
        "유사 사례 및 실적"
    ]

    payload = {
        "company": {
            "company_id": c_id,
            "name": c_name,
            "patent_matching_related": valid_patent_matching_related,
            "has_patent_matching_related": len(valid_patent_matching_related) > 0,
            "purpose": c_purpose,
            "keyword": c_keyword,

            "region": company_region,
            "conduct_list_project": conduct_list_project,
            "has_conduct_project": has_conduct_project,

            "벤처기업여부": company_row.get("벤처기업여부", ""),
            "이노비즈여부": company_row.get("이노비즈여부", ""),
            "메인비즈여부": company_row.get("메인비즈여부", ""),
            "ASTI 여부": company_row.get("ASTI 여부", ""),
            "특구 여부": company_row.get("특구 여부", ""),

            "매출성장율_상위비율": company_row.get("매출성장율_상위비율", None),
            "매출성장율_group": extract_group(company_row.get("매출성장율_판정정보")),

            "영업이익율_상위비율": company_row.get("영업이익율_상위비율", None),
            "영업이익율_group": extract_group(company_row.get("영업이익율_판정정보")),

            "부채비율_하위비율": company_row.get("부채비율_하위비율", None),
            "부채비율_group": extract_group(company_row.get("부채비율_판정정보")),

            "연구개발비_상위비율": company_row.get("연구개발비_상위비율", None),
            "연구개발비_group": extract_group(company_row.get("연구개발비_판정정보")),

        },
        "project": {
            "project_id": pid,
            "title": pname,
            "project_score": pscore,
            "cosine_distance": dist,
            "keyword": keyword_proj,

            "총연구비_상위비율": rec_row.get("총연구비_상위비율", None),
            "총연구비_group": extract_group(rec_row.get("총연구비_판정정보")),
        },

        "conduct_list_company": conduct_list_company,
        "has_conduct_company": has_conduct_company,

        "region_relation": region_relation,

        "related_research": {
            "paper": paper_list,
            "patent": patent_list
        },
        "has_paper": has_paper,
        "has_patent": has_patent,
        "paper_list_count": paper_list_count,
        "patent_list_count": patent_list_count,

        "output_requirements": {
            "language": "ko",
            "format": fmt,
            "section_sentence_range": "각 섹션 4~6문장",
            "forbidden": ["예시", "참고", "제출", "메타"]
        }
    }
    payload = to_py(payload)

    user_prompt = (
        "아래 JSON만을 근거로 추천 근거를 작성해.\n"
        "다만 기술 설명이나 원리 설명이 필요한 경우에는 일반적인 산업 또는 기술 지식을 활용하여 설명할 수 있다.\n"
        "few-shot 예시에 나온 모든 고유명사는 예시 전용이며, 현재 출력에 절대 재사용하지 마.\n"
        "출력은 JSON 하나만 반환해. JSON 외 텍스트 금지야.\n"
        "첫 글자는 {, 마지막 글자는 } 로 끝내.\n"
        "``` 같은 코드블록은 절대 쓰지 마.\n"
        "키는 output_requirements.format에 있는 섹션 제목을 그대로 사용해.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}"
    )


    messages = (
    [{"role": "system", "content": SYSTEM_PROMPT}]
    + FEWSHOT_MESSAGES
    + [{"role": "user", "content": user_prompt}]
    )

    return messages, fmt


def load_model(
    model_id="Qwen/Qwen3.5-35B-A3B-GPTQ-Int4",#"./models/Qwen3.5-35B-A3B-GPTQ-Int4",
    hf_token=None,
    tensor_parallel_size=1,
):
    """
    vLLM 모델과 tokenizer를 로드합니다.

    HF_TOKEN은 코드에 직접 쓰지 말고, 실행 전에 환경변수로 설정하는 것을 권장합니다.
    예:
        export HF_TOKEN="hf_xxx"
    """
    os.environ.setdefault("VLLM_USE_V1", "0")
    os.environ.setdefault("VLLM_WORKER_MULTIPROC_METHOD", "spawn")

    if hf_token is None:
        hf_token = os.environ.get("HF_TOKEN")

    #print('LLM 모델 로딩중 . . .')
    #t_model_start = time.perf_counter()

    llm = LLM(
        model=model_id,
        hf_token=hf_token,
        trust_remote_code=True,
        tensor_parallel_size=tensor_parallel_size,
    )

    tokenizer = AutoTokenizer.from_pretrained(
        model_id,
        trust_remote_code=True,
        token=hf_token,
    )

    #t_model_end = time.perf_counter()
    #elapsed = t_model_end - t_model_start
    #print(f"[완료] LLM 모델 로딩: {format_time(elapsed)}")

    return llm, tokenizer


def _messages_to_prompt(messages):
    """
    vLLM 프롬프트 문자열로 변환합니다.
    assistant 시작 토큰을 넣어줘야 모델이 답변 생성을 시작합니다.
    """
    parts = []

    for m in messages:
        role = m["role"]
        content = m["content"].strip()

        if role == "system":
            parts.append(f"<|im_start|>system\n{content}<|im_end|>")
        elif role == "user":
            parts.append(f"<|im_start|>user\n{content}<|im_end|>")
        elif role == "assistant":
            parts.append(f"<|im_start|>assistant\n{content}<|im_end|>")

    parts.append("<|im_start|>assistant\n")
    return "\n".join(parts)


@torch.inference_mode()
def generate_explanation(
    messages,
    tokenizer,
    model,
    max_new_tokens=4096,
    temperature=0.3,
    top_p=0.9,
):
    prompt = _messages_to_prompt(messages)

    params = SamplingParams(
        temperature=temperature,
        top_p=top_p,
        max_tokens=max_new_tokens,
        stop=["<|im_end|>"],
    )

    outputs = model.generate([prompt], params)
    result = outputs[0].outputs[0].text.strip()

    result = re.sub(r"<think>.*?</think>\s*", "", result, flags=re.DOTALL).strip()
    result = re.sub(r"^\s*</think>\s*", "", result).strip()

    return result



def run_generation(
    tmp,
    model,
    tokenizer,
    out_path="./result/result.csv",
    token_limit=30000,
    max_new_tokens=4096,
    overwrite=True,
):
    
    # 입력 토큰 수 확인
    def count_prompt_tokens(messages, tok):
        prompt = _messages_to_prompt(messages)
        # add_special_tokens=False: 이미 <|im_start|> 같은 토큰을 직접 넣고 있어서 중복 방지
        ids = tok(prompt, add_special_tokens=False).input_ids
        return len(ids)

    """
    project_id 단위로 묶어서 기업별 추천 근거를 생성하고 CSV로 저장합니다.
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if overwrite and out_path.exists():
        out_path.unlink()
    file_exists = out_path.exists()


    for company_id, block in tmp.groupby("company_id", sort=False):
        company_name = block["company_name"].iloc[0]
        company_row = block.iloc[0]
        explanations = []
        company_rows = []

        for _, rec_row in block.iterrows():
            messages, expected_keys = build_company_single_prompt(
                company_row,
                rec_row,
            )

            prompt_tokens = count_prompt_tokens(messages, tokenizer)
            if prompt_tokens > token_limit:
                continue

            one = generate_explanation(
                messages=messages,
                tokenizer=tokenizer,
                model=model,
                max_new_tokens=max_new_tokens,
            )

            company_rows.append({
                "company_id": company_id,
                "company_name": company_name,
                "project_id": rec_row.get("project_id", ""),
                "project_name": rec_row.get("project_name", ""),
                "explanation": one,
            })

        pd.DataFrame(company_rows).to_csv(
            out_path,
            mode="a",
            header=not file_exists,
            index=False,
            encoding="utf-8-sig",
        )
        file_exists = True

        gc.collect()
        if torch.cuda.is_available():
            torch.cuda.empty_cache()


    return {
        "out_path": str(out_path)
    }

'''
def run_llm_pipeline(
    tmp,
    description_path=None,
    model_id="Qwen/Qwen3.5-35B-A3B-GPTQ-Int4",
    hf_token=None,
    tensor_parallel_size=1,
    out_path="./result/result.csv"
):
    print('LLM 생성 시작')
    t_total_start = time.perf_counter()

    model, tokenizer = load_model(
        model_id=model_id,
        hf_token=hf_token,
        tensor_parallel_size=tensor_parallel_size,
    )

    t_total_end = time.perf_counter()

    elapsed = t_total_end - t_total_start
    print(f"[완료] LLM 생성 완료: {format_time(elapsed)}")

    return run_generation(
        tmp=tmp,
        model=model,
        tokenizer=tokenizer,
        out_path=out_path,
    )

    return result
'''