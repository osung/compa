# -*- coding: utf-8 -*-
"""
matching_viewer 추천 근거 생성 모듈 (LLM 백엔드 자동 선택)

- 실행 환경에 맞춰 LLM 백엔드를 자동으로 선택한다 (BACKEND, _detect_backend 참고):
    * Apple Silicon (macOS arm64)      → MLX-LM (mlx-community/Qwen3.5-35B-A3B-4bit)
    * CUDA GPU 사용 가능 (주로 Linux)  → vLLM   (기본 Qwen/Qwen3-8B)
  환경변수 MV_LLM_BACKEND / MV_VLLM_MODEL 등으로 강제·조정할 수 있다.
- 원본의 SYSTEM_PROMPT / FEWSHOT_MESSAGES / 정규화 함수들은
  company_to_project/llm_pipeline.py 에서 그대로 재사용한다.
- matching_viewer 가 보유한 컬럼만으로 payload를 구성하므로,
  원본 pipeline에 비해 누락된 필드는 빈 값으로 채워진다 (MISSING_FIELDS 참고).
"""

import json
import os
import pickle
import re
import sys
import threading
import types

import numpy as np
import pandas as pd

# company_to_project/llm_pipeline.py 는 최상단에서 vLLM 을 import 한다.
# - CUDA 환경처럼 vLLM 이 실제로 설치돼 있으면 그대로 진짜 모듈을 쓴다(우리 백엔드가 사용).
# - macOS(Apple Silicon)처럼 vLLM 이 없으면, 그 파일의 프롬프트/헬퍼만 재사용하고
#   vLLM 의존 코드(load_model, generate_explanation)는 호출하지 않으므로,
#   import 통과만을 위해 dummy 모듈을 등록한다.
import importlib.util as _ilu  # noqa: E402

if "vllm" not in sys.modules and _ilu.find_spec("vllm") is None:
    _stub = types.ModuleType("vllm")
    _stub.LLM = object  # 본 모듈에서는 호출하지 않음
    _stub.SamplingParams = object
    sys.modules["vllm"] = _stub

# 두 방향의 프롬프트를 모두 사용한다:
#  - company_to_project: 기업→과제 추천 ('추천 과제의 우수성')
#  - project_to_company: 과제→기업 추천 ('추천 기업 우수성')
# 두 디렉터리에 같은 이름(llm_pipeline.py)이 있어 sys.path 로는 하나만 잡히므로,
# 파일 경로로 각각 명시적으로 로드한다. 정규화 헬퍼는 양쪽이 동일하므로 c2p 것을 쓴다.
import importlib.util  # noqa: E402

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))


def _load_pipeline(modname, subdir):
    path = os.path.join(_THIS_DIR, subdir, "llm_pipeline.py")
    spec = importlib.util.spec_from_file_location(modname, path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[modname] = mod
    spec.loader.exec_module(mod)
    return mod


_c2p = _load_pipeline("mv_pipeline_c2p", "company_to_project")
_p2c = _load_pipeline("mv_pipeline_p2c", "project_to_company")

# 공통 정규화 헬퍼 (양쪽 동일 → c2p 것 사용)
normalize_str_list = _c2p.normalize_str_list
normalize_conduct_list_company = _c2p.normalize_conduct_list_company
normalize_conduct_list_project = _c2p.normalize_conduct_list_project
analyze_region_relation = _c2p.analyze_region_relation
to_py = _c2p.to_py

# 방향별 SYSTEM_PROMPT / FEWSHOT / 섹션 포맷
#   direction='company' : 기업에서 유사 과제 추천 → '추천 과제의 우수성'
#   direction='project' : 과제에서 유사 기업 추천 → '추천 기업 우수성'
_DIRECTIONS = {
    "company": {
        "system": _c2p.SYSTEM_PROMPT,
        "fewshot": _c2p.FEWSHOT_MESSAGES,
        "fmt": ["연관성", "추천 과제의 우수성", "유사 사례 및 실적"],
    },
    "project": {
        "system": _p2c.SYSTEM_PROMPT,
        "fewshot": _p2c.FEWSHOT_MESSAGES,
        "fmt": ["연관성", "추천 기업 우수성", "유사 사례 및 실적"],
    },
}

# 하위호환: 기존 코드가 참조하던 기본(기업→과제) 프롬프트 심볼
SYSTEM_PROMPT = _c2p.SYSTEM_PROMPT
FEWSHOT_MESSAGES = _c2p.FEWSHOT_MESSAGES


def _resolve_direction(direction):
    return direction if direction in _DIRECTIONS else "company"


# 원본 pipeline 이 기대하지만 matching_viewer 데이터(+보강 데이터셋)에도 여전히 없는 필드 목록.
# UI에서 사용자에게 안내하기 위한 정보용. 누락된 필드는 빈 값으로 채운다.
# (df_company_dataset / df_project_dataset 조인으로 논문·특허 성과, 총연구비 등급,
#  재무지표, 인증, 보유특허는 보강되었고, 아래 항목만 데이터 부재로 남아 있다.)
MISSING_FIELDS = {
    "company": [
        "연구개발비 (상위비율 + group 판정정보 — 회사 데이터셋에 없음)",
    ],
    "project": [
        # conduct_list_company 는 과제 수행기업(사업자등록번호)이 기업 universe 에
        # 있을 때 채워지므로, 항상 누락은 아니다(universe 밖이면 빈 값).
    ],
    "note": (
        "보강 데이터셋 조인으로 과제 논문·특허 성과, 총연구비 등급, 기업 재무지표·인증·보유특허, "
        "사업목적, 과제 수행기업(conduct_list_company), 회사 수행과제 이력(conduct_list_project)이 "
        "채워진다. 보유특허가 10개 이상인 기업은 과제 임베딩과의 유사도 top-3 특허를, 과제의 논문·"
        "특허 성과가 10개 이상이면 기업 임베딩과의 유사도 top-3 를 선별해 사용하고(pro-sroberta), "
        "10개 미만이면 전체를 사용한다. 다만 연구개발비 지표는 데이터셋에 없어 사용되지 않으며, "
        "수행기업/수행과제가 데이터셋에 없는 경우 해당 근거가 비어 있을 수 있다."
    ),
}


# ----------------------------------------------------------------------
# LLM 백엔드 선택
#   - Apple Silicon (macOS arm64)      → MLX-LM (mlx-community/Qwen3.5-35B-A3B-4bit)
#   - CUDA GPU 사용 가능 (주로 Linux)  → vLLM   (기본 Qwen/Qwen3.5-35B-A3B-GPTQ-Int4)
#   환경변수로 강제/조정 가능:
#     MV_LLM_BACKEND = mlx | vllm        (자동 감지 무시)
#     MV_VLLM_MODEL  = HF repo id        (vLLM 모델, 기본 Qwen/Qwen3.5-35B-A3B-GPTQ-Int4)
#     MV_VLLM_GPU_UTIL / MV_VLLM_MAX_LEN / MV_VLLM_DTYPE  (vLLM 메모리 튜닝)
#
#   vLLM 기본값은 MoE 모델 Qwen/Qwen3.5-35B-A3B-GPTQ-Int4
#   (총 35B / 활성 3B, GPTQ Int4) 이다. 활성 파라미터가
#   3B 라 dense 8B 대비 토큰 생성이 빠르고, Int4 양자화로 가중치 용량이 작아
#   통합 메모리 안에 가중치+KV 가 충분히 들어간다.
# ----------------------------------------------------------------------
MLX_MODEL_ID = os.environ.get("MV_MLX_MODEL", "mlx-community/Qwen3.5-35B-A3B-4bit")
VLLM_MODEL_ID = os.environ.get("MV_VLLM_MODEL", "Qwen/Qwen3.5-35B-A3B-GPTQ-Int4")


def _detect_backend():
    """실행 환경에 맞는 LLM 백엔드('mlx' 또는 'vllm')를 결정."""
    forced = os.environ.get("MV_LLM_BACKEND", "").strip().lower()
    if forced in ("mlx", "vllm"):
        return forced
    import platform

    # Apple Silicon → MLX
    if sys.platform == "darwin" and platform.machine() == "arm64":
        return "mlx"
    # CUDA 사용 가능 → vLLM
    try:
        import torch

        if torch.cuda.is_available():
            return "vllm"
    except Exception:
        pass
    # 최종 fallback: macOS면 mlx, 그 외엔 vllm
    return "mlx" if sys.platform == "darwin" else "vllm"


BACKEND = _detect_backend()
MODEL_ID = MLX_MODEL_ID if BACKEND == "mlx" else VLLM_MODEL_ID


# ======================================================================
# 보강 데이터셋 (논문·특허 성과 / 재무·인증 / 보유특허) 참조
# ----------------------------------------------------------------------
# matching_viewer 가 로드하는 *_with_desc.pkl 에는 기업·과제 '설명문'과 키워드만
# 있고, 추천 근거의 '추천 과제의 우수성'·'유사 사례' 섹션을 채우는 데 필요한
#   - 과제: 논문명_리스트 / 특허명_리스트 / 총연구비 상위비율
#   - 기업: 보유특허 / 재무지표(매출성장율·영업이익율·부채비율) / 인증(벤처·이노비즈·메인비즈)
# 은 별도 통합 데이터셋에만 존재한다. 아래 두 파일을 과제고유번호 / 사업자번호로
# 조인해 payload 를 보강한다. (설명문은 기존대로 *_with_desc.pkl 의 것을 사용한다.)
# 파일이 없거나 로드 실패 시에는 보강을 건너뛰고 기존(빈 값) 동작으로 fallback 한다.
# ======================================================================
PROJECT_DATASET_FILE = os.path.join(_THIS_DIR, "df_project_dataset_260602_with_ranks.pkl")
COMPANY_DATASET_FILE = os.path.join(_THIS_DIR, "df_company_dataset_260604_with_projects.pkl")

# 원본 데이터셋은 pandas 3.0(+pyarrow/numpy2)로 저장돼 구버전 pandas 에서는 로드되지
# 않는다. 그래서 convert_datasets_for_matching_viewer.py 로 '필요한 컬럼만 추출한
# plain-python dict' 슬림 파일을 만들어 두면(.lookup.pkl) 어떤 pandas/numpy 에서도
# 즉시 로드된다. 슬림 파일이 있으면 그것을 우선 사용하고, 없으면 원본을 직접 읽는다.
PROJECT_LOOKUP_FILE = os.path.join(_THIS_DIR, "df_project_dataset_260602_with_ranks.lookup.pkl")
COMPANY_LOOKUP_FILE = os.path.join(_THIS_DIR, "df_company_dataset_260604_with_projects.lookup.pkl")

# 논문·특허가 매우 많은 과제에서 프롬프트 입력 토큰 폭증·컨텍스트 초과를 막기 위한 상한.
# payload 에 직렬화되는 '제목 수'만 제한하며, *_count 에는 전체 개수를 보존한다.
MAX_PAPER_TITLES = 30
MAX_PATENT_TITLES = 30
MAX_COMPANY_PATENT_TITLES = 15
MAX_CONDUCT_PROJECTS = 15        # 회사 수행과제 이력 중 추천 과제와 관련도 높은 상위 N개
MAX_PURPOSE_ITEMS = 15           # 사업목적 항목 상한

# 조인에 필요한 컬럼만 추려 메모리에 적재한다 (임베딩벡터 등 대용량 컬럼은 제외).
_PROJECT_DATASET_COLS = [
    "논문명_리스트", "특허명_리스트",
    "총연구비_상위비율", "총연구비_판정정보",
    "사업자등록번호",   # 과제 수행기업 → conduct_list_company 조인 키
    "과제수행기관명",   # 결과 테이블 표시용(수행기관)
    "논문건수", "특허건수",                         # 결과 테이블 표시용
    "논문건수_상위비율", "논문건수_판정정보",        # 분야 내 상위비율(우수성)
    "특허건수_상위비율", "특허건수_판정정보",
]
_COMPANY_DATASET_COLS = [
    "업체명", "키워드_리스트", "사업목적",   # conduct_list_company·purpose 구성용
    "보유특허명_리스트", "특허건수", "지역", "중분류명",
    "벤처기업여부", "이노비즈여부", "메인비즈여부", "ASTI여부", "특구여부",
    "매출성장율_상위비율", "영업이익율_상위비율", "부채비율_하위비율",
    "수행과제고유번호", "수행과제명",        # conduct_list_project(회사 수행 과제 이력) 구성용
]

_project_lookup = None    # {과제고유번호(str): {col: value}}  / 로드 실패 시 {}
_company_lookup = None     # {사업자번호(str): {col: value}}
_dataset_lock = threading.Lock()


def _norm_id(v):
    """ID 정규화: 공백 제거 + float 직렬화('123.0') 흔적 제거."""
    s = _safe_str(v)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def _build_dataset_lookup(path, id_col, cols):
    """대용량 통합 데이터셋(pkl)을 id_col 기준 dict 로 인덱싱.

    임베딩벡터 등 불필요한 대용량 컬럼은 버리고 조인에 필요한 컬럼만 남긴다.
    파일이 없거나 로드 실패 시 빈 dict 를 반환해 호출부가 보강 없이 동작하도록 한다.
    """
    if not os.path.exists(path):
        return {}
    try:
        df = pd.read_pickle(path)
    except Exception:
        return {}
    if id_col not in df.columns:
        return {}
    keep = [id_col] + [c for c in cols if c in df.columns]
    df = df[keep].copy()
    lookup = {}
    for rec in df.to_dict("records"):
        lookup[_norm_id(rec.get(id_col))] = rec
    del df
    lookup.pop("", None)
    return lookup


def _load_slim_lookup(path):
    """변환된 슬림 lookup 파일(plain-python dict pickle)을 로드. 실패 시 None."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "rb") as f:
            obj = pickle.load(f)
        return obj if isinstance(obj, dict) else None
    except Exception:
        return None


def _get_project_dataset():
    global _project_lookup
    if _project_lookup is None:
        with _dataset_lock:
            if _project_lookup is None:
                _project_lookup = (
                    _load_slim_lookup(PROJECT_LOOKUP_FILE)
                    or _build_dataset_lookup(
                        PROJECT_DATASET_FILE, "과제고유번호", _PROJECT_DATASET_COLS)
                )
    return _project_lookup


def _get_company_dataset():
    global _company_lookup
    if _company_lookup is None:
        with _dataset_lock:
            if _company_lookup is None:
                _company_lookup = (
                    _load_slim_lookup(COMPANY_LOOKUP_FILE)
                    or _build_dataset_lookup(
                        COMPANY_DATASET_FILE, "사업자번호", _COMPANY_DATASET_COLS)
                )
    return _company_lookup


def _extract_group(info):
    """총연구비_판정정보({'rank':int,'group':str}) 에서 group 문자열을 추출."""
    if isinstance(info, dict):
        return _safe_str(info.get("group"))
    return ""


def _cap_titles(items, n):
    """제목 리스트를 상위 n개로 제한. (전체 개수는 호출부에서 별도 보존)"""
    items = items or []
    return items[:n] if len(items) > n else items


def _ratio_int(v):
    """상위/하위 비율(1/5/10/20) 값을 int 로 정규화. 유효하지 않으면 None."""
    try:
        if v is None or pd.isna(v):
            return None
    except (TypeError, ValueError):
        return None
    try:
        iv = int(v)
    except (TypeError, ValueError):
        return None
    return iv if iv in (1, 5, 10, 20) else None


def _yn(v):
    """'Y'/'N' 플래그 정규화. yes 계열이면 'Y', no 계열이면 'N', 그 외 ''."""
    s = _safe_str(v).upper()
    if s in ("Y", "YES", "1", "TRUE", "T"):
        return "Y"
    if s in ("N", "NO", "0", "FALSE", "F"):
        return "N"
    return ""


def _build_conduct_company(perf_bizno, exclude_bizno=None):
    """과제 수행기업(사업자등록번호) → conduct_list_company 한 건 구성.

    수행기업을 기업 데이터셋(df_company_dataset)과 조인해 company_info(이름·지역·
    키워드·사업목적·보유특허·수행과제)를 채운다. 기업 universe 에 없거나 사업자번호가
    유효하지 않으면([] / '0' 등) 서술 근거가 될 수 없으므로 빈 리스트를 반환한다.

    exclude_bizno(매칭/추천 기업의 사업자번호)와 수행기업이 동일하면 '자기 자신'이라
    유사 사례가 될 수 없으므로 제외한다.
    """
    biz = _norm_id(perf_bizno)
    if not biz or biz == "0" or len(biz) < 10:
        return []
    if exclude_bizno is not None and biz == _norm_id(exclude_bizno):
        return []   # 과제수행기관 == 매칭 기업 → 자기 자신이므로 제외
    perf = _get_company_dataset().get(biz)
    if not perf:
        return []   # universe 밖: 기업명/정보를 채울 수 없어 제외
    name = _safe_str(perf.get("업체명"))
    info = {
        "company_name": name,
        "region": _safe_str(perf.get("지역")),
        "company_keyword": normalize_str_list(perf.get("키워드_리스트", [])),
        "company_purpose_list": normalize_str_list(perf.get("사업목적", [])),
        "company_patent": _cap_titles(
            normalize_str_list(perf.get("보유특허명_리스트", [])),
            MAX_COMPANY_PATENT_TITLES),
        # 수행기관이 수행한 과제(협업 필터링: 매칭 기업의 수행과제와 비교)
        "conduct_projects": _cap_titles(
            normalize_str_list(perf.get("수행과제명", [])),
            MAX_CONDUCT_PROJECTS),
    }
    # 이름·키워드·특허·수행과제 등 서술 가능한 정보가 하나도 없으면 제외
    if not (name or info["company_keyword"] or info["company_patent"]
            or info["company_purpose_list"] or info["conduct_projects"]):
        return []
    return [{
        "company_id": biz,
        "company_name": name,
        "company_info": info,
    }]


def _simple_tokens(text):
    """간단 토큰화(한글/영숫자 2글자 이상). 수행과제 관련도 랭킹용."""
    return set(t for t in re.split(r"[^0-9A-Za-z가-힣]+", str(text)) if len(t) >= 2)


# ---- 보유특허 top-3 유사도 선별 (pro-sroberta) — lazy 로드 ----
_patent_topk = {"loaded": False, "build_fn": None, "model": None, "noun_fn": None}
_patent_topk_lock = threading.Lock()


def _get_patent_topk_tools():
    """(build_company_patent_sim, pro-sroberta model, noun_fn) 를 lazy 로 준비.

    company_to_project/data_pipeline.py 의 검증된 top-3 로직을 그대로 재사용한다.
    로드 실패(모델/의존성 부재) 시 (None, None, None) → 호출부가 전체목록으로 폴백.
    최초 '특허 10개 이상' 기업의 근거 생성 때 1회만 로드된다.
    """
    if _patent_topk["loaded"]:
        return _patent_topk["build_fn"], _patent_topk["model"], _patent_topk["noun_fn"]
    with _patent_topk_lock:
        if not _patent_topk["loaded"]:
            try:
                path = os.path.join(_THIS_DIR, "company_to_project", "data_pipeline.py")
                spec = importlib.util.spec_from_file_location("mv_dp_c2p", path)
                dp = importlib.util.module_from_spec(spec)
                sys.modules["mv_dp_c2p"] = dp
                spec.loader.exec_module(dp)
                model, _, _ = dp.load_embedding_model("pro-sroberta")
                _patent_topk["build_fn"] = dp.build_company_patent_sim
                _patent_topk["model"] = model
                _patent_topk["noun_fn"] = dp._get_default_noun_fn()
            except Exception:
                pass
            _patent_topk["loaded"] = True
    return _patent_topk["build_fn"], _patent_topk["model"], _patent_topk["noun_fn"]


def _select_by_similarity(items, ref_embed, cap, desc=None):
    """제목 리스트를 기준 임베딩(ref_embed, pro-sroberta)과의 cosine 유사도로 선별.

    - 10개 이상 → 명사 임베딩 cosine top-3 (build_company_patent_sim 재사용, threshold 0.5)
    - 10개 미만이거나 임베딩 불가/유사도 미달 → 상한(cap) 내 전체로 폴백

    기업 보유특허(기준=과제 임베딩)와 과제 논문·특허 성과(기준=기업 임베딩) 양쪽에
    동일하게 사용한다.
    """
    if not items:
        return []
    if len(items) >= 10 and ref_embed is not None:
        build_fn, model, noun_fn = _get_patent_topk_tools()
        if build_fn is not None and model is not None:
            try:
                sel = build_fn(items, ref_embed, model,
                               threshold=0.5, top_n=3, min_for_topk=10,
                               embed_mode="noun", noun_fn=noun_fn,
                               show_progress=True, progress_desc=desc)
                if sel:
                    return sel
            except Exception:
                pass
    return _cap_titles(items, cap)


def _select_company_patents(company_patents, project_row):
    """보유특허 → patent_matching_related (기준 임베딩 = 매칭 과제 임베딩)."""
    project_embed = project_row.get("norm_embed") if hasattr(project_row, "get") else None
    return _select_by_similarity(company_patents, project_embed, MAX_COMPANY_PATENT_TITLES,
                                 desc="기업 보유특허 유사 top-3")


def _build_conduct_project(c_extra, ref_text, cap=None):
    """회사가 수행한 과제(수행과제고유번호/수행과제명) → conduct_list_project.

    수가 많을 수 있어(수천 건) 추천 과제와 이름 토큰이 많이 겹치는 순으로 상위
    cap개만 남긴다(임베딩이 없어 토큰 겹침으로 근사). 과제명만 있으므로
    project_info 의 keyword/paper/patent 는 비워 둔다(프롬프트는 과제명으로 비교).
    """
    if cap is None:
        cap = MAX_CONDUCT_PROJECTS
    ids = c_extra.get("수행과제고유번호") or []
    names = c_extra.get("수행과제명") or []
    pairs = [(_norm_id(i), _safe_str(n)) for i, n in zip(ids, names)]
    pairs = [(i, n) for i, n in pairs if i or n]
    if not pairs:
        return []
    ref = _simple_tokens(ref_text)
    if ref:
        pairs.sort(key=lambda pn: len(ref & _simple_tokens(pn[1])), reverse=True)
    out = []
    for pid, pname in pairs[:cap]:
        out.append({
            "project_id": pid,
            "project_name": pname,
            "project_info": {
                "project_name": pname,
                "project_keyword": [],
                "paper": [],
                "patent": [],
            },
        })
    return out


# ======================================================================
# 사용자 추가 입력 → 키워드 추출 (matching_viewer.py 와 동일 규칙 이식)
# ----------------------------------------------------------------------
# matching_viewer.py 는 유사 검색 시 사용자가 입력한 추가 데이터
#  - 기업측: 주요 제품(main_products) / 사업목적(business_purpose)
#  - 과제측: 한글 키워드(korean_keywords) / 과학기술표준분류(sci_tech_class)
# 에서 _normalize_user_input 으로 '명사 키워드'를 추출해 임베딩 입력 텍스트에
# 병합한다. 추천 근거(LLM payload)도 동일한 추가 입력을 반영해야 임베딩 매칭
# 근거와 설명이 일치하므로, 같은 불용어/접미사/정제 로직을 그대로 가져온다.
# (gen_compound_keywords.py 의 STOPWORDS/JOSA_EOMI 와 동일 어휘)
# ======================================================================
KEYWORD_SEP = ";"

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
USER_INPUT_SUFFIXES = (
    "판매업", "도매업", "소매업", "서비스업", "제조업", "생산업",
    "가공업", "임대업", "대행업", "무역업", "유통업", "공업", "업",
)


def _is_meaningless_code(token):
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


def _normalize_user_input(text):
    """사용자 자유입력(주요 제품/사업목적/한글 키워드/과기표준분류)을 명사 키워드로 정제.

    업종 문장체("반도체 제품 도매업")를 그대로 넣으면 행정/정책 문체로 끌려가므로
    핵심 명사만 추출한다. matching_viewer.MatchingViewer._normalize_user_input 와 동일.
    절차:
    1) 구두점/구분자를 공백으로 치환 후 어절 단위로 분리
    2) 2글자 미만·불용어 제거
    3) 의미 없는 코드(과기표준분류 코드 'EH030301' 등) 제거
    4) 업종 접미사(~판매업/~도매업/~업 등) 제거 후 남는 핵심 명사 재검사
    반환: 정제된 키워드 토큰 리스트(등장 순서 유지).
    """
    if not text:
        return []
    text = re.sub(r'[^\w\s가-힣a-zA-Z0-9]', ' ', str(text))
    out = []
    for w in text.split():
        if len(w) < 2 or w.casefold() in USER_INPUT_STOPWORDS:
            continue
        if _is_meaningless_code(w):
            continue
        stem = w
        for suf in USER_INPUT_SUFFIXES:
            if stem.endswith(suf) and len(stem) > len(suf):
                stem = stem[:-len(suf)]
                break
        stem = stem.strip()
        if len(stem) < 2 or stem.casefold() in USER_INPUT_STOPWORDS:
            continue
        out.append(stem)
    return out


def _merge_user_keywords(base_keywords, user_texts, fallback_text=None):
    """기존 키워드 리스트에 사용자 추가입력에서 추출한 명사 키워드를 병합.

    matching_viewer._build_combined_text / _build_combined_text_project 와 동일한
    병합 규칙을 따르되, 임베딩용 문자열(';' join) 대신 payload용 키워드 리스트를 반환한다.
    - base_keywords: 행의 기존 키워드 리스트(이미 normalize_str_list 통과한 값)
    - user_texts: 사용자 자유입력 문자열들 (예: [주요제품, 사업목적])
    - fallback_text: 기존 키워드가 3개 이하일 때만 통째로 덧붙일 보조 텍스트
                     (기업=10차산업코드명, 과제=과제명)
    중복은 strip + casefold 정규화 키로 제거하고, 먼저 등장한 원형을 유지한다.
    """
    parts = []
    seen = set()

    def add(token):
        if token is None:
            return
        token = str(token).strip()
        key = token.casefold()
        if not key or key in seen:
            return
        seen.add(key)
        parts.append(token)

    base_len = 0
    if base_keywords:
        for k in base_keywords:
            if k is None:
                continue
            base_len += 1
            add(k)

    for text in (user_texts or []):
        if not text:
            continue
        for tok in _normalize_user_input(text):
            add(tok)

    if base_len <= 3 and fallback_text:
        fb = _safe_str(fallback_text)
        if fb:
            add(fb)

    return parts


# -------- MLX-LM 모델 캐시 (프로세스 1회 로드) ---------
_model = None
_tokenizer = None
_model_lock = threading.Lock()
_model_load_error = None


def _safe_str(v):
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return "" if s.lower() in {"nan", "none", "null"} else s


def is_model_ready():
    return _model is not None


def _load_model_mlx(progress_cb=None):
    if progress_cb:
        progress_cb(f"MLX 모델 로드 중 ({MODEL_ID})...")
    from mlx_lm import load as mlx_load

    return mlx_load(MODEL_ID)


def _load_model_vllm(progress_cb=None):
    if progress_cb:
        progress_cb(f"vLLM 모델 로드 중 ({MODEL_ID})...")
    # JSON 스키마 강제(StructuredOutputsParams)가 검증된 조합(system 파이프라인)과 맞도록
    # V0 엔진 + spawn 워커로 고정. 필요시 환경변수로 덮어쓸 수 있다.
    os.environ.setdefault("VLLM_USE_V1", "0")
    os.environ.setdefault("VLLM_WORKER_MULTIPROC_METHOD", "spawn")
    from vllm import LLM
    from transformers import AutoTokenizer

    # 메모리/분산 파라미터를 환경변수로 조정 가능하게 둔다.
    gpu_util = float(os.environ.get("MV_VLLM_GPU_UTIL", "0.90"))
    max_len = int(os.environ.get("MV_VLLM_MAX_LEN", "8192"))
    dtype = os.environ.get("MV_VLLM_DTYPE", "auto")
    tp_size = int(os.environ.get("MV_VLLM_TP", "1"))   # 멀티 GPU 분산(tensor parallel)

    llm = LLM(
        model=MODEL_ID,
        trust_remote_code=True,
        dtype=dtype,
        gpu_memory_utilization=gpu_util,
        max_model_len=max_len,
        max_num_seqs=int(os.environ.get("MV_VLLM_MAX_SEQS", "1")),
        tensor_parallel_size=tp_size,
        enforce_eager=True,  # CUDA 그래프 캡처 메모리 절약
    )
    tokenizer = AutoTokenizer.from_pretrained(MODEL_ID, trust_remote_code=True)
    return llm, tokenizer


def load_model_blocking(progress_cb=None):
    """LLM 모델을 동기적으로 로드. 첫 호출 시 디스크에서 메모리로 적재.

    BACKEND('mlx'|'vllm') 에 따라 적절한 백엔드로 로드한다.
    progress_cb(msg): 진행 상태 콜백 (UI 라벨 갱신용).
    """
    global _model, _tokenizer, _model_load_error
    with _model_lock:
        if _model is not None:
            return _model, _tokenizer
        if _model_load_error is not None:
            raise _model_load_error
        try:
            if BACKEND == "mlx":
                _model, _tokenizer = _load_model_mlx(progress_cb)
            else:
                _model, _tokenizer = _load_model_vllm(progress_cb)
            if progress_cb:
                progress_cb(f"{BACKEND.upper()} 모델 로드 완료")
            return _model, _tokenizer
        except Exception as e:
            _model_load_error = e
            raise


def build_payload(
    company_row,
    project_row,
    match_score=None,
    similarity=None,
    promise=None,
    company_user_data=None,
    project_user_data=None,
    direction="company",
):
    """matching_viewer 에서 보유한 정보만으로 LLM payload(JSON) 구성.

    company_row: pandas Series — company_embeddings_*.pkl의 한 행
        예상 컬럼: 사업자번호, 한글업체명, 10차산업코드, 10차산업코드명,
                   키워드_리스트, 유망성점수, ASTI 여부, 특구 여부, 설립일
    project_row: pandas Series — public_RnD_embeddings_*.pkl의 한 행
        예상 컬럼: 과제고유번호, 과제명, 제출년도, 키워드_리스트, 유망성점수

    company_user_data: dict | None — 유사 검색 시 사용자가 기업에 추가 입력한 데이터.
        {'main_products': 주요 제품, 'business_purpose': 사업목적}
        matching_viewer 가 임베딩 재생성에 쓴 입력과 동일. 여기서 _normalize_user_input
        으로 명사 키워드를 추출해 회사 키워드/사업목적에 병합한다.
    project_user_data: dict | None — 사용자가 과제에 추가 입력한 데이터.
        {'korean_keywords': 한글 키워드, 'sci_tech_class': 과학기술표준분류}
        과제 키워드에 병합한다.

    추가 입력을 반영하면 임베딩 매칭 근거(재생성된 벡터)와 LLM 추천 근거가
    같은 키워드 집합 위에서 작성되어 설명 일관성이 유지된다.
    """
    company_user_data = company_user_data or {}
    project_user_data = project_user_data or {}

    # ---- 회사 ----
    c_id = _safe_str(company_row.get("사업자번호"))
    c_name = _safe_str(company_row.get("한글업체명"))
    industry = _safe_str(company_row.get("10차산업코드명"))

    # 보강 데이터셋(df_company_dataset)에서 재무·인증·보유특허 등 추가 컬럼 조인
    c_extra = _get_company_dataset().get(_norm_id(c_id), {})

    # 기존 키워드 리스트 + 사용자 추가입력(주요 제품/사업목적)에서 추출한 명사 키워드
    main_products = company_user_data.get("main_products", "")
    business_purpose = company_user_data.get("business_purpose", "")
    c_keyword = _merge_user_keywords(
        normalize_str_list(company_row.get("키워드_리스트", [])),
        [main_products, business_purpose],
        fallback_text=industry,
    )

    # 기업 설명문(LLM 생성) + 품질 플래그 — 추천 근거 작성에 핵심 컨텍스트로 사용
    c_desc = _safe_str(company_row.get("기업설명문"))
    c_desc_issue = _safe_str(company_row.get("desc_issue"))
    c_desc_ok = company_row.get("desc_ok") if hasattr(company_row, "get") else None
    try:
        c_desc_ok = int(c_desc_ok) if c_desc_ok is not None and not pd.isna(c_desc_ok) else None
    except (TypeError, ValueError):
        c_desc_ok = None

    # 사업목적(purpose): 데이터셋의 사업목적 + 사용자 입력 사업목적을 함께 사용하고,
    # 업종명을 보조로 덧붙인다. (사용자 입력은 명사 키워드로 정제해 합침)
    c_purpose = []
    seen_purpose = set()

    def _add_purpose(item):
        s = _safe_str(item)
        key = s.casefold()
        if s and key not in seen_purpose:
            seen_purpose.add(key)
            c_purpose.append(s)

    for item in normalize_str_list(c_extra.get("사업목적", [])):   # 데이터셋 사업목적
        _add_purpose(item)
    for tok in (_normalize_user_input(business_purpose) if business_purpose else []):
        _add_purpose(tok)                                         # 사용자 입력 사업목적
    if industry:
        _add_purpose(industry)                                    # 업종명 보조
    c_purpose = c_purpose[:MAX_PURPOSE_ITEMS]

    # 인증/구분: 보강 데이터셋 우선, 없으면 임베딩 행의 값 사용
    asti = _yn(c_extra.get("ASTI여부")) or _safe_str(company_row.get("ASTI 여부"))
    teukgu = _yn(c_extra.get("특구여부")) or _safe_str(company_row.get("특구 여부"))
    venture = _yn(c_extra.get("벤처기업여부"))
    innobiz = _yn(c_extra.get("이노비즈여부"))
    mainbiz = _yn(c_extra.get("메인비즈여부"))

    # 회사 보유특허(보유특허명_리스트) → patent_matching_related.
    # 특허가 10개 이상이면 과제 임베딩(project_row.norm_embed, pro-sroberta)과의
    # cosine 유사도 top-3 를 뽑고(build_company_patent_sim 재사용), 10개 미만이면 전체.
    # 임베딩 모델 로드/임베딩이 불가하면 상한 내 전체로 안전 폴백한다.
    company_patents = normalize_str_list(c_extra.get("보유특허명_리스트", []))
    selected_patents = _select_company_patents(company_patents, project_row)
    valid_patent_matching_related = [{"특허명": t} for t in selected_patents]
    # 보유특허 건수 — '추천 기업 우수성'에서 기업 기술역량 근거로 사용
    try:
        _pc = c_extra.get("특허건수")
        company_patent_count = (
            int(_pc) if _pc is not None and not pd.isna(_pc) else None)
    except (TypeError, ValueError):
        company_patent_count = None

    # conduct_list_project(회사가 수행한 과제 이력)와 conduct_list_company(추천 과제를
    # 수행한 기업)는 추천 과제 정보가 필요하므로 아래 과제 섹션에서 채운다.

    # 재무지표(같은 업종 내 상위/하위 비율). group 은 읽기 쉬운 중분류명(없으면 업종명).
    rev_growth = _ratio_int(c_extra.get("매출성장율_상위비율"))
    op_margin = _ratio_int(c_extra.get("영업이익율_상위비율"))
    debt_ratio = _ratio_int(c_extra.get("부채비율_하위비율"))
    fin_group = _safe_str(c_extra.get("중분류명")) or industry or None

    # ---- 과제 ----
    p_id = _safe_str(project_row.get("과제고유번호"))
    p_name = _safe_str(project_row.get("과제명"))

    # 기존 키워드 리스트 + 사용자 추가입력(한글 키워드/과학기술표준분류)에서 추출한 명사 키워드
    korean_keywords = project_user_data.get("korean_keywords", "")
    sci_tech_class = project_user_data.get("sci_tech_class", "")
    p_keyword = _merge_user_keywords(
        normalize_str_list(project_row.get("키워드_리스트", [])),
        [korean_keywords, sci_tech_class],
        fallback_text=p_name,
    )

    # 과제 설명문(LLM 생성) — 추천 근거 작성에 핵심 컨텍스트로 사용
    p_desc = _safe_str(project_row.get("과제설명문"))

    # 보강 데이터셋(df_project_dataset)에서 논문·특허 성과 / 총연구비 등급 조인
    p_extra = _get_project_dataset().get(_norm_id(p_id), {})
    paper_all = normalize_str_list(p_extra.get("논문명_리스트", []))
    patent_all = normalize_str_list(p_extra.get("특허명_리스트", []))
    paper_list_count = len(paper_all)
    patent_list_count = len(patent_all)
    # 과제 논문·특허 성과가 10개 이상이면 기업 임베딩(pro-sroberta)과 cosine 유사도
    # top-3 를 선별(기업 보유특허와 동일 로직), 10개 미만이면 상한 내 전체.
    # count(*_list_count)는 전체 개수를 보존해 '논문 N건' 같은 정량 서술이 정확하다.
    company_embed = company_row.get("norm_embed") if hasattr(company_row, "get") else None
    paper_list = _select_by_similarity(paper_all, company_embed, MAX_PAPER_TITLES,
                                       desc="과제 논문성과 유사 top-3")
    patent_list = _select_by_similarity(patent_all, company_embed, MAX_PATENT_TITLES,
                                        desc="과제 특허성과 유사 top-3")
    has_paper = paper_list_count > 0
    has_patent = patent_list_count > 0

    # 총연구비 상위비율(1/5/10/20) + group(과학기술표준분류1-중)
    fund_ratio = _ratio_int(p_extra.get("총연구비_상위비율"))
    fund_group = _extract_group(p_extra.get("총연구비_판정정보"))

    # 분야(과학기술표준분류1-중) 내 논문·특허 건수 상위비율(1/5/10/20) + group
    paper_rank = _ratio_int(p_extra.get("논문건수_상위비율"))
    paper_rank_group = _extract_group(p_extra.get("논문건수_판정정보"))
    patent_rank = _ratio_int(p_extra.get("특허건수_상위비율"))
    patent_rank_group = _extract_group(p_extra.get("특허건수_판정정보"))

    # 추천 과제를 수행한 기업 → conduct_list_company ('유사 사례'의 수행기업·지역 근거)
    # 과제 데이터셋의 사업자등록번호(수행기업)를 기업 데이터셋과 조인해 company_info 구성.
    # 기업 universe 에 없어 정보를 채울 수 없으면 서술 근거가 못 되므로 제외한다.
    conduct_list_company = _build_conduct_company(
        p_extra.get("사업자등록번호"), exclude_bizno=c_id)

    # 회사가 수행한 과제 이력 → conduct_list_project ('유사 사례'의 기업 수행과제 근거).
    # 추천 과제명/키워드와 이름이 겹치는 순으로 상위 N개만 사용(과다 방지).
    conduct_list_project = _build_conduct_project(
        c_extra, ref_text=f"{p_name} {';'.join(p_keyword)}")

    # cosine_distance = 1 - similarity (LLM 프롬프트에서 참고용)
    if similarity is not None:
        cosine_distance = float(max(0.0, 1.0 - float(similarity)))
    else:
        cosine_distance = ""

    project_score = (
        float(match_score) if match_score is not None else
        _safe_str(project_row.get("유망성점수"))
    )

    fmt = list(_DIRECTIONS[_resolve_direction(direction)]["fmt"])

    company_region = _safe_str(c_extra.get("지역"))
    region_relation = analyze_region_relation(company_region, conduct_list_company)

    payload = {
        "company": {
            "company_id": c_id,
            "name": c_name,
            "description": c_desc,
            "desc_issue": c_desc_issue,
            "desc_ok": c_desc_ok,
            "patent_matching_related": valid_patent_matching_related,
            "has_patent_matching_related": len(valid_patent_matching_related) > 0,
            "patent_count": company_patent_count,
            # 보유특허 목록 — '연관성' 섹션(프롬프트 1-3)에서 회사 기술축 파악에 사용
            "company_patent": _cap_titles(company_patents, MAX_COMPANY_PATENT_TITLES),
            "purpose": c_purpose,
            "keyword": c_keyword,
            "region": company_region,
            "conduct_list_project": conduct_list_project,
            "has_conduct_project": len(conduct_list_project) > 0,

            "벤처기업여부": venture,
            "이노비즈여부": innobiz,
            "메인비즈여부": mainbiz,
            "ASTI 여부": asti,
            "특구 여부": teukgu,

            "매출성장율_상위비율": rev_growth,
            "매출성장율_group": fin_group if rev_growth else None,
            "영업이익율_상위비율": op_margin,
            "영업이익율_group": fin_group if op_margin else None,
            "부채비율_하위비율": debt_ratio,
            "부채비율_group": fin_group if debt_ratio else None,
            "연구개발비_상위비율": None,   # 데이터셋에 없음
            "연구개발비_group": None,
        },
        "project": {
            "project_id": p_id,
            "title": p_name,
            "description": p_desc,
            "project_score": project_score,
            "cosine_distance": cosine_distance,
            "keyword": p_keyword,
            "총연구비_상위비율": fund_ratio,
            "총연구비_group": fund_group if fund_ratio else None,
            "논문건수_상위비율": paper_rank,
            "논문건수_group": paper_rank_group if paper_rank else None,
            "특허건수_상위비율": patent_rank,
            "특허건수_group": patent_rank_group if patent_rank else None,
        },
        "conduct_list_company": conduct_list_company,
        "has_conduct_company": len(conduct_list_company) > 0,
        "region_relation": region_relation,
        "related_research": {
            "paper": paper_list,
            "patent": patent_list,
        },
        "has_paper": has_paper,
        "has_patent": has_patent,
        "paper_list_count": paper_list_count,
        "patent_list_count": patent_list_count,
        "output_requirements": {
            "language": "ko",
            "format": fmt,
            "section_sentence_range": "각 섹션 4~6문장",
            "forbidden": ["예시", "참고", "제출", "메타"],
        },
    }
    return to_py(payload), fmt


def build_messages(payload, direction="company"):
    d = _DIRECTIONS[_resolve_direction(direction)]
    user_prompt = (
        "아래 JSON만을 근거로 추천 근거를 작성해.\n"
        "다만 기술 설명이나 원리 설명이 필요한 경우에는 일반적인 산업 또는 기술 지식을 활용하여 설명할 수 있다.\n"
        "company.description(기업 설명문)과 project.description(과제 설명문)은 각각 기업과 과제를 요약한 "
        "핵심 컨텍스트다. 키워드(keyword)만이 아니라 이 설명문을 반드시 읽고, 기업의 사업 영역·강점과 과제의 "
        "목표·내용을 연결지어 '연관성' 섹션의 근거로 적극 활용해.\n"
        "단, company.desc_ok 가 0 이면 기업 설명문에 품질 이슈(company.desc_issue)가 있다는 뜻이므로 "
        "그 내용은 보조 참고로만 쓰고 keyword 등 다른 정보를 우선해. desc_ok 가 1 이면 설명문을 신뢰해서 사용해도 된다.\n"
        "설명문 문장을 그대로 복사하지 말고, 추천 맥락에 맞게 재구성해서 서술해.\n"
        "few-shot 예시에 나온 모든 고유명사는 예시 전용이며, 현재 출력에 절대 재사용하지 마.\n"
        "출력은 JSON 하나만 반환해. JSON 외 텍스트 금지야.\n"
        "첫 글자는 {, 마지막 글자는 } 로 끝내.\n"
        "``` 같은 코드블록은 절대 쓰지 마.\n"
        "키는 output_requirements.format에 있는 섹션 제목을 그대로 사용해.\n\n"
        f"{json.dumps(payload, ensure_ascii=False)}"
    )

    return (
        [{"role": "system", "content": d["system"]}]
        + d["fewshot"]
        + [{"role": "user", "content": user_prompt}]
    )


def _strip_think(text):
    """모델이 가끔 출력하는 <think>...</think> 블록 제거"""
    text = re.sub(r"<think>.*?</think>\s*", "", text, flags=re.DOTALL)
    text = re.sub(r"^\s*</think>\s*", "", text)
    return text.strip()


def _build_prompt(tokenizer, messages):
    """Qwen3 instruct 템플릿. thinking 모드는 가능하면 끈다."""
    try:
        return tokenizer.apply_chat_template(
            messages,
            tokenize=False,
            add_generation_prompt=True,
            enable_thinking=False,
        )
    except TypeError:
        # 구버전 템플릿: enable_thinking 미지원
        return tokenizer.apply_chat_template(
            messages,
            tokenize=False,
            add_generation_prompt=True,
        )


def _stream_explanation_mlx(
    messages, *, max_tokens, temperature, top_p, on_token, should_stop, expected_keys=None
):
    # MLX-LM 은 JSON 스키마 강제(grammar)를 지원하지 않으므로 expected_keys 는 생성 단계에서
    # 사용하지 않는다. 대신 temperature=0.0(결정적) + 프롬프트 + parse_sections 의 키 강제
    # (expected_keys 로만 구성)로 출력 구조를 실질 보장한다.
    model, tokenizer = load_model_blocking()
    from mlx_lm.generate import stream_generate
    from mlx_lm.sample_utils import make_sampler

    prompt = _build_prompt(tokenizer, messages)
    sampler = make_sampler(temp=temperature, top_p=top_p)

    pieces = []
    for resp in stream_generate(
        model,
        tokenizer,
        prompt=prompt,
        max_tokens=max_tokens,
        sampler=sampler,
    ):
        chunk = resp.text or ""
        pieces.append(chunk)
        if on_token is not None and chunk:
            on_token(chunk)
        if should_stop is not None and should_stop():
            break

    return _strip_think("".join(pieces))


def _json_object_schema(expected_keys):
    """섹션 키들을 'string 값 필수 + 추가키 불가' JSON object 스키마로 구성."""
    keys = list(expected_keys)
    return {
        "type": "object",
        "properties": {k: {"type": "string"} for k in keys},
        "required": keys,
        "additionalProperties": False,
    }


def _vllm_structured_kwargs(expected_keys):
    """vLLM SamplingParams 에 넘길 JSON 스키마 강제 인자를 버전별로 구성.

    - 신버전: StructuredOutputsParams(json=...)  (system 파이프라인과 동일)
    - 구버전: GuidedDecodingParams(json=...)
    미지원/미설치면 {} → 자유 생성으로 폴백.
    """
    if not expected_keys:
        return {}
    schema = _json_object_schema(expected_keys)
    try:
        from vllm.sampling_params import StructuredOutputsParams
        return {"structured_outputs": StructuredOutputsParams(json=schema)}
    except Exception:
        pass
    try:
        from vllm.sampling_params import GuidedDecodingParams
        return {"guided_decoding": GuidedDecodingParams(json=schema)}
    except Exception:
        pass
    return {}


def _stream_explanation_vllm(
    messages, *, max_tokens, temperature, top_p, on_token, should_stop, expected_keys=None
):
    # vLLM 오프라인 LLM.generate 는 배치(비스트리밍)이므로, 생성 완료 후
    # 전체 텍스트를 on_token 으로 한 번에 전달한다.
    # expected_keys 가 주어지면 JSON 스키마를 강제해 항상 유효 JSON(필수 키만)을 보장한다.
    model, tokenizer = load_model_blocking()
    from vllm import SamplingParams

    prompt = _build_prompt(tokenizer, messages)
    params = SamplingParams(
        temperature=temperature,
        top_p=top_p,
        max_tokens=max_tokens,
        stop=["<|im_end|>"],
        **_vllm_structured_kwargs(expected_keys),
    )
    outputs = model.generate([prompt], params)
    text = outputs[0].outputs[0].text or ""
    if on_token is not None and text:
        on_token(text)
    return _strip_think(text)


def stream_explanation(
    messages,
    *,
    max_tokens=2048,
    temperature=0.0,
    top_p=1.0,
    on_token=None,
    should_stop=None,
    expected_keys=None,
):
    """추천 근거를 생성. on_token(text_chunk) 콜백으로 (부분) 텍스트 전달.

    - MLX  : 토큰 단위 스트리밍 (JSON 스키마 강제 미지원 → temperature=0.0 + parse_sections 키 강제).
    - vLLM : 배치 생성 후 전체 텍스트를 한 번에 콜백 (expected_keys 시 JSON 스키마 강제).
    temperature 기본값은 0.0(결정적). expected_keys=섹션 제목 목록(fmt) 을 넘기면 출력 구조를 보장.
    반환: 전체 생성 텍스트 (think 블록 제거 후)
    """
    fn = _stream_explanation_mlx if BACKEND == "mlx" else _stream_explanation_vllm
    return fn(
        messages,
        max_tokens=max_tokens,
        temperature=temperature,
        top_p=top_p,
        on_token=on_token,
        should_stop=should_stop,
        expected_keys=expected_keys,
    )


# 수행 기업·지역 연관성 등 '정보 부재/생략' 자체를 서술하는 문장 제거용 패턴.
# 프롬프트로 1차 억제하되, 모델이 가끔 흘리는 부재 서술을 후처리로 확실히 걷어낸다.
_ABSENCE_PAT = re.compile(
    r"(제공되지\s*않|정보[가는은]\s*없|해당\s*(내용|정보|사례)\S*\s*생략|"
    r"생략(합니다|한다|됩니다|됨|하)|분석할\s*수\s*없|작성할\s*수\s*없|활용되지\s*않|"
    r"수행\s*기업[이가]?\s*(존재하지\s*않|없)|존재하지\s*않|확인되지\s*않|conduct_list_company)"
)


def _scrub_absence(text):
    """'정보가 없어 생략한다' 류의 부재·생략 메타 서술 문장을 제거한다.
    한국어 종결('다.'/'요.') 기준으로 문장을 나눠 해당 패턴이 든 문장만 버린다."""
    if not text or not _ABSENCE_PAT.search(text):
        return text
    sents = re.split(r"(?<=[다요]\.)", text)
    kept = "".join(s for s in sents if not _ABSENCE_PAT.search(s)).strip()
    return kept if kept else text


def _section_to_text(v):
    """섹션 값을 깔끔한 문자열로 정규화.

    모델이 섹션 값을 '문장 리스트'로 반환하면 str(list) 가 "['..','..']" 처럼
    '[' 로 시작해 ']' 로 끝나는 텍스트가 된다. 이를 막기 위해 리스트는 이어붙이고,
    통째로 대괄호로 감싼 문자열은 (JSON 배열이면 파싱해서) 풀어준다.
    마지막으로 부재·생략 메타 서술 문장을 제거(_scrub_absence)한다.
    """
    if isinstance(v, (list, tuple)):
        return _scrub_absence(" ".join(_section_to_text(x) for x in v).strip())
    s = str(v).strip()
    if len(s) >= 2 and s[0] == "[" and s[-1] == "]":
        try:
            arr = json.loads(s)
            if isinstance(arr, list):
                return _scrub_absence(" ".join(_section_to_text(x) for x in arr).strip())
        except Exception:
            pass
        s = s[1:-1].strip()       # 파싱 실패 시 양끝 대괄호만 제거
    return _scrub_absence(s)


def parse_sections(text, fmt=("연관성", "추천 과제의 우수성", "유사 사례 및 실적")):
    """모델이 반환한 JSON(또는 JSON 유사) 문자열에서 섹션별 텍스트를 추출.

    JSON 파싱 실패 시에는 단순 정규식으로 섹션 헤더를 잘라 fallback.
    """
    text = text.strip()
    # 가장 흔한 케이스: 첫 '{' ~ 마지막 '}' 구간을 JSON 으로 시도
    try:
        first = text.find("{")
        last = text.rfind("}")
        if first >= 0 and last > first:
            obj = json.loads(text[first : last + 1])
            if isinstance(obj, dict):
                return {k: _section_to_text(obj.get(k, "")) for k in fmt}
    except Exception:
        pass

    # Fallback: 섹션 헤더 기반 분할
    out = {k: "" for k in fmt}
    pattern = "|".join(re.escape(k) for k in fmt)
    parts = re.split(rf"\b({pattern})\b\s*[:：]?\s*", text)
    # parts: ['', sec1, body1, sec2, body2, ...]
    for i in range(1, len(parts), 2):
        sec = parts[i]
        body = parts[i + 1] if i + 1 < len(parts) else ""
        if sec in out:
            out[sec] = _section_to_text(body)
    return out
