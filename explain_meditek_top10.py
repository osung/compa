# -*- coding: utf-8 -*-
"""MEDITEK TOP10 매칭 결과 → 보고서 입력 JSON(상세 추천 근거 + 표용 매칭 근거 생성).

match_meditek_top10.py 는 순위 선정까지만 하므로(적합도·우수성), 보고서에 필요한 서술은
여기서 35B 로 생성한다. 텍스트는 모두 모델이 쓰고, 이 스크립트는 프롬프트 구성과
체크포인트 관리만 한다.

  · 추천근거_상세 : 3섹션(연관성 · 기술 적합성 · 추천 과제의 우수성).
                   보유기술이 있으면 '기술 적합성'을 보강·고도화 관점으로, 없으면 수요 충족
                   관점으로 쓰게 가이드만 바꾼다 — 섹션 제목은 같게 두어 보고서에 매칭 기준
                   (수요기술/기보유기술)이 드러나지 않게 한다.
  · 판단근거     : 표에 싣는 긍정형 한 문장(적합도 채점 단계의 비판적 reason 을 쓰지 않는다).

수요기술은 기업이 확보하려는(미보유) 기술, 기보유기술은 이미 보유한 역량이므로
payload 에서 각각 company.수요기술_* / company.description 으로 분리해 넣는다
(둘을 섞으면 '이미 보유했다'는 서술 오류가 난다).

특허 실적은 _patent_prep_nice.py 가 만든 pid_patents.json(= 보고서에 싣는 것과 같은 소스)을
쓰고, 유망성 점수는 프롬프트·산출 어디에도 넣지 않는다(assert_no_promise 로 매번 검사).

사용: COMPA_SCRATCH=<scratch> python explain_meditek_top10.py [--tag MEDITEK] [--only 3,7]
산출: MEDITEK_TOP10_보고서.json (+ top10_<tag>_{explain,reason}_ckpt.json)
"""
import argparse
import json
import os
import re
import time

import pandas as pd

import compa_match as cm
import match_meditek_top10 as mt

TOP10_PKL = "MEDITEK_TOP10.pkl"
OUT_JSON = "MEDITEK_TOP10_보고서.json"
RETRY = 2                  # 금지 표현(매칭 기준 노출)이 섞였을 때 재생성 횟수

# ---- 상세 근거 섹션 구성 ----------------------------------------------------
# 섹션 제목은 두 경우 모두 동일하게 둔다 — 보고서에 '수요기술/기보유기술' 중 무엇을
# 기준으로 매칭했는지가 드러나지 않아야 한다(관점 차이는 가이드 문구로만 준다).
# '유사 사례 및 실적' 은 뺐다 — 실측 130행에서 앞 세 섹션과 어휘 62.6% 중복, 문장 완전
# 중복 20%, 새 성과를 제시한 행은 18%뿐이었다. 지시문('과제의 논문·특허 실적이 수요
# 해결에 주는 시사점')이 '추천 과제의 우수성'('연구성과(논문/특허)…강점')과 겹치고,
# 특허·논문 실적 표를 바로 아래에 싣게 되어 나열의 정보 가치도 사라졌다.
EX_FMT = ["연관성", "기술 적합성", "추천 과제의 우수성"]

_G_COMMON = {
    # 유망성 점수는 제공하지 않으므로 근거로 삼지 않는다(점수·등급 언급 금지).
    # 독자가 검증할 수 있도록 주어진 건수는 반드시 숫자로 인용하게 한다(비율 표현만 쓰면
    # '상위 10% 수준' 같은 추상적 평가만 남아 근거의 확인 가능성이 떨어진다).
    "추천 과제의 우수성": ("과제의 연구성과(논문·특허)·수행기관 역량·연구 규모 등 추천 과제의 강점. "
                   "project.patent_list_count(특허 건수)와 paper_list_count(논문 건수)로 주어진 "
                   "수치는 반드시 숫자로 인용해 서술한다(예: '특허 6건, 논문 3건의 성과를 확보'). "
                   "건수가 0이면 인용하지 말고 다른 강점을 서술한다. 상위비율이 주어지면 함께 쓰되, "
                   "주어지지 않은 점수·등급·유망성 수치는 언급하지 말 것"),
    "유사 사례 및 실적": cm._EX_GUIDE["유사 사례 및 실적"],
}
_G_RELATION = ("company.description(이 기업이 '이미 보유한' 기술·사업내용)과 company.수요기술명·"
               "수요기술_내용(이 기업이 '확보하려는', 아직 보유하지 않은 기술)을 구분해 읽고, "
               "이 기업의 기술적 위치에 비추어 과제의 목표·내용이 어떤 지점에서 맞닿는지 설명. "
               "수요기술을 기업이 이미 보유한 역량으로 서술하지 말 것. 반대로 description 에 있는 "
               "보유기술은 기업이 실제로 가진 것으로 서술해도 된다. 둘 중 한쪽이 비어 있으면 "
               "있는 쪽만 근거로 쓰고 없는 역량을 임의로 가정하지 말 것")

# 확보하려는 기술만 주어진 기업 / 보유 기술이 함께 주어진 기업 — 같은 섹션명, 다른 관점
_GUIDE_DEM = dict(_G_COMMON, **{
    "연관성": _G_RELATION,
    "기술 적합성": ("과제가 company.수요기술_내용의 요구를 어떤 방식으로 충족·해결할 수 있는지, "
              "과제의 기술 요소와 수요 항목을 대응시켜 구체적으로 설명. 완전히 일치하지 "
              "않는 항목은 어느 범위까지 기여할 수 있는지로 서술"),
})
_GUIDE_HOLD = dict(_G_COMMON, **{
    "연관성": _G_RELATION,
    "기술 적합성": ("과제의 기술이 company.description 의 보유기술을 어떻게 보강·고도화하거나 "
              "새로운 응용으로 확장할 수 있는지 설명. 보유기술의 어느 구성요소(소재·공정·"
              "알고리즘·계측 등)에 결합되는지, 그 결합이 성능·적용범위·신뢰성 측면에서 "
              "무엇을 개선하는지 인과적으로 서술"),
})

# ---- 매칭 기준 비노출 -------------------------------------------------------
# 보고서에는 무엇을 기준으로 매칭했는지(확보 희망 기술 / 이미 보유한 기술)가 드러나면 안 된다.
# 프롬프트에서 금지어를 명시하고, 위반하면 재생성하며, 끝까지 남으면 기계적으로 치환한다.
BANNED = re.compile(r"수요\s?기술|기보유\s?기술|기술\s?수요|보유\s?기술|확보하려는|"
                    r"수요\s?충족|보유\s?보강|수요\s?요구|수요와|수요의|수요에")
_BAN_RULE = ("**매칭 기준을 가리키는 표현을 쓰지 마라. '수요기술'·'기보유기술'·'기술수요'·"
             "'보유 기술'·'확보하려는'·'수요 충족'·'수요 요구' 같은 말은 단 하나도 쓰지 말고, "
             "기업의 기술을 가리킬 때는 '당사 기술', '이 기업의 기술', '<기업명>의 기술' 처럼 "
             "서술한다.**")
# 최종 치환(문장 성립을 유지하는 최소 치환). 모델 출력에 남은 위반과, 기업 제출 원문에
# 라벨로 박혀 있는 '수요기술명/수요기술 개요' 같은 표기에도 적용한다. 라벨성 표기를 먼저.
_FIX = [(re.compile(r"수요\s?기술\s?명"), "기술명"),
        (re.compile(r"수요\s?기술\s?개요"), "기술 개요"),
        (re.compile(r"수요\s?기술\s?내용"), "기술 내용"),
        (re.compile(r"수요\s?기술\s?사양"), "기술 사양"),
        (re.compile(r"확보하려는\s*"), ""), (re.compile(r"기보유\s?기술"), "당사 기술"),
        (re.compile(r"수요\s?기술"), "당사 기술"), (re.compile(r"기술\s?수요"), "당사 기술"),
        (re.compile(r"보유\s?기술"), "당사 기술"), (re.compile(r"수요\s?충족"), "기술 부합"),
        (re.compile(r"보유\s?보강"), "기술 보강"), (re.compile(r"수요\s?요구"), "기술 요건"),
        (re.compile(r"수요와"), "당사 기술과"), (re.compile(r"수요의"), "당사 기술의"),
        (re.compile(r"수요에"), "당사 기술에")]


def debanned(text):
    """금지 표현 최종 치환 → (정리된 텍스트, 치환 여부)."""
    out = text
    for pat, rep in _FIX:
        out = pat.sub(rep, out)
    return out, out != text


def log(*a):
    print(*a, flush=True)


# 모델 비교용 계측(지시 준수·포맷 안정성). 생성 로직에는 영향 없음.
STATS = {"상세_호출": 0, "상세_재시도": 0, "상세_최종위반": 0, "상세_섹션부족": 0,
         "근거_호출": 0, "근거_재시도": 0, "근거_최종위반": 0}


# ---- 표기 정리 --------------------------------------------------------------
# 원본 hwp 에서 넘어온 '3 D 배양'·'( PDC)' 같은 잔재가 모델 출력에도 옮겨 붙는다.
# 숫자-단위 사이 공백만 붙이되, 조사/구두점이 뒤따르는 경우로 한정해 '3 개발' → '3개발'
# 같은 오수정을 막는다.
_JOSA = r"(?=[의을를이가와과로에도만은는,·\s.)\]]|$)"
_NUM_SAFE = re.compile(r"(\d)\s+(차원|개월|시간|%|D(?![A-Za-z가-힣])|nm|㎛|mm|cm)")
_NUM_JOSA = re.compile(r"(\d)\s+(건|년|명|개|종|배|차)" + _JOSA)
_PAREN_L = re.compile(r"\(\s+")
_PAREN_R = re.compile(r"\s+\)")


_PUBLIC_RND = re.compile(r"공공\s*R&D")     # 명칭 통일: 공공 R&D → 국가 R&D


def polish(s):
    """숫자-단위·괄호 공백 잔재 정리 + 마침표 뒤 공백 보정 + 매칭 기준 표현/명칭 정리."""
    s = str(s or "")
    s = _PUBLIC_RND.sub(lambda m: m.group(0).replace("공공", "국가"), s)
    s = _PAREN_L.sub("(", _PAREN_R.sub(")", s))
    s = _NUM_SAFE.sub(r"\1\2", s)
    s = _NUM_JOSA.sub(r"\1\2", s)
    s, _ = debanned(s)
    return cm.normalize_spacing(s)


# ---- 유망성 점수 차단 --------------------------------------------------------
# 유망성 점수는 프롬프트에도 보고서에도 들어가면 안 된다. payload 를 만들 때마다 검사한다.
_PROMISE_KEY = re.compile(r"유망|promise|project_score", re.I)


def assert_no_promise(obj, where="payload"):
    """dict/list 를 훑어 유망성 관련 키가 값(숫자·문자)을 갖고 있으면 중단."""
    def walk(o, path):
        if isinstance(o, dict):
            for k, v in o.items():
                if _PROMISE_KEY.search(str(k)) and v not in ("", None, [], {}):
                    raise SystemExit(f"[중단] {where}: 유망성 정보 유입 {path}/{k} = {v!r}")
                walk(v, f"{path}/{k}")
        elif isinstance(o, (list, tuple)):
            for i, v in enumerate(o):
                walk(v, f"{path}[{i}]")
    walk(obj, "")
    return obj


def hold_body(unit):
    """기보유기술 합본에서 [기술명] 섹션을 뺀 본문(기술명은 별도 필드로 넣는다)."""
    return mt._HOLD_NAME_SEC.sub("", unit["기보유기술 내용"]).strip()


def payload_for(unit, kws, proj):
    demand = {
        "기업명": unit["기업명"],
        "수요기술명": unit["수요기술명"],
        "수요기술 내용": unit["수요기술 내용"],
        "수요기술 사양": "",
        "예상 적용 제품 및 서비스": "",
        "keywords": kws,
    }
    body = hold_body(unit)
    if body:                                   # 보유기술 = 기업의 실제 역량 → description
        demand["기업설명문"] = body[:2200]
        demand["desc_ok"] = 1
        demand["desc_issue"] = ""
    p = cm.build_demand_payload(demand, proj)
    guide = _GUIDE_HOLD if body else _GUIDE_DEM
    if body:
        p["company"]["보유기술명"] = unit["기보유기술명"]
        p["company"]["보유기술_구분"] = " / ".join(
            x for x in (unit["기술유형"], unit["기술분야"]) if x)
    p["project"]["project_score"] = ""            # 유망성 점수 미제공
    p["output_requirements"].update({
        "format": EX_FMT,
        "section_guide": {k: guide[k] for k in EX_FMT},
        "must_cover_수요기술_사양": False,
        "금지_표현": ["수요기술", "기보유기술", "기술수요", "보유 기술", "확보하려는",
                  "수요 충족", "수요 요구"],
        "기업_기술_지칭": "당사 기술 / 이 기업의 기술 / <기업명>의 기술",
    })
    return assert_no_promise(p, "상세근거 프롬프트"), EX_FMT


def gen_detail(unit, kws, proj, retry=RETRY):
    """3섹션 상세 근거 → 한 셀 텍스트. 금지 표현이 섞이면 다시 생성한다."""
    p, fmt = payload_for(unit, kws, proj)
    msgs = cm.build_messages(p, direction="company")
    msgs[-1]["content"] += "\n" + _BAN_RULE
    text = ""
    STATS["상세_호출"] += 1
    for i in range(retry + 1):
        out = cm.stream_explanation(msgs, max_tokens=1400, temperature=0.2 if i == 0 else 0.0,
                                    top_p=0.9, expected_keys=fmt)
        secs = cm.parse_sections(out, tuple(fmt))
        parts = [f"[{k}] {secs.get(k, '').strip()}" for k in fmt if secs.get(k, "").strip()]
        if len(parts) < len(fmt):
            STATS["상세_섹션부족"] += 1
        text = cm.normalize_spacing("\n\n".join(parts) if parts else out.strip())
        if not BANNED.search(text):
            return text
        STATS["상세_재시도"] += 1
    STATS["상세_최종위반"] += 1
    return text                                  # 남은 위반은 JSON 조립 단계에서 치환


# 표용 한 문장 근거 — cm._MR_SYS/_MR_FEWSHOT 은 '수요와 부합' 류 표현을 예시로 쓰므로
# (매칭 기준이 드러난다) 이 보고서용으로 규칙과 예시를 따로 둔다. 문장은 모델이 생성한다.
_MR_SYS = (
    "너는 기업의 기술과 국가 R&D 과제의 '매칭 근거'를 한 문장으로 요약하는 한국어 AI다. "
    "이 과제가 왜 해당 기업에 추천되는지, 두 대상이 공유하는 핵심 기술·목적을 근거로 긍정적으로 서술한다.\n"
    "작성 규칙:\n"
    "1) 60자 이내 한 문장. 명사형/음슴체 종결(예: '~ 기술이 당사 기술과 부합', '~에 활용 가능').\n"
    "2) 과제 개요·수행기관·기간·분야 등 '설명'은 쓰지 말고, 기업 기술과 과제의 '접점'만 쓴다.\n"
    "3) 부정 평가('~ 부재', '~ 미포함', '~ 불일치', '~ 미흡', '다름', '부족')는 쓰지 않는다. "
    "완전히 일치하지 않아도 공유하는 기술적 접점을 중심으로 '~에 활용 가능', '~ 기반 마련', "
    "'~ 부분 부합' 처럼 기여 가능성으로 표현한다.\n"
    "4) 문장 하나만 출력. 따옴표·머리기호·부연 금지.\n"
    "5) few-shot 예시는 형식·톤 참고용이며, 예시의 고유명사·문구를 재사용하지 말고 현재 입력만 근거로 작성한다.\n"
    "6) " + _BAN_RULE)


def _mr_demo(tech, proj, ans):
    return [{"role": "user", "content": f"[기업 기술]\n{tech}\n\n[R&D 과제]\n{proj}\n\n"
                                        "매칭 근거(한 문장):"},
            {"role": "assistant", "content": ans}]


_MR_FEWSHOT = (
    _mr_demo("기술명: 도라지 사포닌 정제·표준화 및 분말 제형화 기술\n"
             "기술 내용: 다년근 도라지 유래 플라티코딘 D 등 사포닌을 표준화하고 분말 제형으로 안정화",
             "과제명: 도라지 사포닌 추출 및 정제 공정 표준화 연구\n"
             "과제설명: 도라지에서 사포닌을 고효율로 추출·정제하고 분말화하는 표준 공정 개발",
             "도라지 사포닌 추출·정제·분말화 표준 공정이 당사 기술과 직접 일치")
    + _mr_demo("기술명: 엣지 컴퓨터 기반 차량 번호판 인식 기술\n"
               "기술 내용: 저전력 엣지 디바이스 실시간 번호판 OCR용 딥러닝 모델 경량화",
               "과제명: 엣지 인공지능 자동화 기술\n"
               "과제설명: 신경망 모델 압축·양자화 및 자동 설계로 엣지 저전력 실시간 추론 구현",
               "모델 압축·양자화 기반 엣지 저전력 실시간 추론 기술이 경량 OCR 고도화에 활용 가능")
    + _mr_demo("기술명: 오가노이드 3D 배양·이미징 통합 자동화 플랫폼\n"
               "기술 내용: 나노리터 정밀 분주와 형광 이미징을 결합한 고속 스크리닝 장비",
               "과제명: 3차원 뇌조직 신경활성도 정밀 측정용 스마트 플레이트 개발\n"
               "과제설명: 고밀도 전극 어레이로 3D 배양 조직의 전기·형광 신호를 동시 계측",
               "고밀도 전극 기반 3D 조직 신호 계측 기술이 당사 플랫폼 고도화에 활용 가능")
)


def gen_reason(unit, proj, retry=RETRY):
    """표 '매칭 근거' 한 문장(긍정형). 금지 표현이 섞이면 다시 생성. 실패 시 빈 문자열."""
    blocks = []
    if unit["수요기술 내용"]:
        blocks.append(f"기술명: {unit['수요기술명']}\n"
                      f"기술 내용: {unit['수요기술 내용'][:600]}")
    body = hold_body(unit)
    if body:
        blocks.append(f"기술명: {unit['기보유기술명']}\n"
                      f"기술 내용(보유): {body[:600]}")
    pj = f"과제명: {proj.get('과제명','')}\n과제설명: {str(proj.get('설명','') or '')[:600]}"
    user = ("[기업 기술]\n" + "\n\n".join(blocks)
            + f"\n\n[R&D 과제]\n{pj}\n\n매칭 근거(한 문장):")
    msgs = ([{"role": "system", "content": _MR_SYS}] + _MR_FEWSHOT
            + [{"role": "user", "content": user}])
    out = ""
    STATS["근거_호출"] += 1
    for _ in range(retry + 1):
        try:
            raw = cm.stream_explanation(msgs, max_tokens=120, temperature=0.0,
                                        top_p=1.0).strip()
        except Exception as e:
            log(f"      ! 매칭근거 실패: {e}")
            return ""
        out = raw.strip().strip('"').strip("'").split("\n")[0].strip()
        if out.startswith("매칭 근거"):
            out = out.split(":", 1)[-1].strip()
        out = cm.normalize_spacing(out[:120])
        if not BANNED.search(out):
            return out
        STATS["근거_재시도"] += 1
    STATS["근거_최종위반"] += 1
    return out                                   # 남은 위반은 JSON 조립 단계에서 치환


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=TOP10_PKL)
    ap.add_argument("--tag", default="MEDITEK")
    ap.add_argument("--out", default=OUT_JSON)
    ap.add_argument("--patents",
                    default=os.path.join(os.environ.get("COMPA_SCRATCH", "."),
                                         "pid_patents.json"),
                    help="특허 실적 JSON(_patent_prep_nice.py 산출)")
    ap.add_argument("--only", default="", help="처리할 번호/관리번호(쉼표 구분)")
    ap.add_argument("--limit", type=int, default=0, help="처리 기업 수 제한(0=전체)")
    ap.add_argument("--no-llm", action="store_true",
                    help="LLM 호출 없이 체크포인트에 있는 근거만으로 JSON 재조립")
    a = ap.parse_args()

    df = pd.read_pickle(a.src)
    df["과제고유번호"] = df["과제고유번호"].astype(str)
    units, _ = mt.load_units()
    units.sort(key=mt.sort_key)
    ubykey = {u["키"]: u for u in units}
    kw_ckpt = json.load(open(os.path.join(cm.OUT_DIR, f"top10_{a.tag}_keywords_ckpt.json"),
                             encoding="utf-8"))

    ex_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_explain_ckpt.json")
    rs_path = os.path.join(cm.OUT_DIR, f"top10_{a.tag}_reason_ckpt.json")
    ex = json.load(open(ex_path, encoding="utf-8")) if os.path.exists(ex_path) else {}
    rs = json.load(open(rs_path, encoding="utf-8")) if os.path.exists(rs_path) else {}
    log(f"체크포인트: 상세근거 {len(ex)}건 · 매칭근거 {len(rs)}건")

    # 과제 메타(논문 실적·상위비율) — 상세근거 프롬프트의 우수성/실적 근거
    import pickle
    with open(cm.PROJECT_META, "rb") as f:
        pmeta = pickle.load(f)
    # 특허 실적은 보고서에 싣는 것과 같은 소스를 쓴다(_patent_prep_nice.py 산출)
    if not os.path.exists(a.patents):
        raise SystemExit(f"없는 파일: {a.patents} — 먼저 `python _patent_prep_nice.py` 실행 필요")
    patents = json.load(open(a.patents, encoding="utf-8"))
    assert_no_promise(patents, "특허 데이터")
    log(f"특허 실적 로드: {a.patents} · 과제 {len(patents)}건 · "
        f"특허 {sum(len(v) for v in patents.values())}건")

    todo = [u for u in units if u["키"] in set(df["기업명"].map(cm.norm_name))]
    if a.only:
        keep = {s.strip() for s in a.only.split(",") if s.strip()}
        todo = [u for u in todo if u["번호"] in keep or u["관리번호"] in keep]
    if a.limit:
        todo = todo[:a.limit]

    need = sum(1 for u in todo
               for _ in df[df["기업명"] == u["기업명"]].itertuples())
    log(f"대상 기업 {len(todo)}건 · 추천 {need}건")
    if not a.no_llm:
        log("· 35B 로드…")
        cm.load_model_blocking(progress_cb=lambda m: log("  " + m))

    best, t0, done = {}, time.time(), 0
    for n, u in enumerate(todo, 1):
        key = u["번호"] or u["관리번호"]
        g = df[df["기업명"] == u["기업명"]].sort_values("rank")
        kws = kw_ckpt.get(u["키"], [])
        tops = []
        for r in g.to_dict("records"):
            pid = str(r["과제고유번호"])
            ck = f"{key}::{pid}"
            meta = pmeta.get(pid, {}) if isinstance(pmeta, dict) else {}
            pats = patents.get(pid, [])           # 특허 실적: 새 특허 데이터 기준
            proj = {
                "pid": pid, "과제명": r["과제명"], "설명": r["과제설명문"],
                # 유망성 점수는 넣지 않는다(build_demand_payload 의 project_score 가 빈 값이 된다)
                "수행기관": r["과제수행기관"],
                "키워드": [], "논문명": meta.get("논문명_리스트") or [],
                "특허명": [x["특허명"] for x in pats],
                "논문건수": int(r["논문건수"]), "특허건수": len(pats),
                "총연구비_상위비율": meta.get("총연구비_상위비율"),
                "논문건수_상위비율": meta.get("논문건수_상위비율"),
            }
            assert_no_promise(proj)
            if not a.no_llm and ck not in ex:
                ex[ck] = gen_detail(u, kws, proj)[:cm.DESC_OUT]
                json.dump(ex, open(ex_path, "w", encoding="utf-8"), ensure_ascii=False)
            if not a.no_llm and ck not in rs:
                rs[ck] = gen_reason(u, proj)
                json.dump(rs, open(rs_path, "w", encoding="utf-8"), ensure_ascii=False)
            done += 1
            tops.append({
                "rank": int(r["rank"]), "과제고유번호": pid, "과제명": r["과제명"],
                "수행기관": r["과제수행기관"],
                "적합도": int(r["적합도"]), "수요충족": int(r["수요충족"]),
                "보유보강": int(r["보유보강"]), "최종점수": float(r["최종점수"]),
                # 특허건수는 보고서에 싣는 특허 실적(새 특허 데이터)과 같은 값이어야 한다
                "특허건수": len(pats), "논문건수": int(r["논문건수"]),
                "순위구분": r["순위구분"],          # 유망성 점수는 싣지 않는다
                # 표에는 긍정형 한 문장, 없으면 채점 단계 근거로 폴백
                "판단근거": polish(rs.get(ck) or r["LLM근거"]),
                "과제설명문": r["과제설명문"],
                "추천근거_상세": polish(ex.get(ck, "")),
            })
        best[key] = {
            "번호": u["번호"], "관리번호": u["관리번호"], "기업명": u["기업명"],
            "소스": u["소스"], "키워드": cm.SEP.join(kws),
            "수요기술명": polish(u["수요기술명"]), "수요기술 내용": polish(u["수요기술 내용"]),
            "기보유기술명": polish(u["기보유기술명"]), "기보유기술 내용": polish(hold_body(u)),
            "기술유형": u["기술유형"], "기술분야": u["기술분야"],
            "top10": tops,
        }
        el = time.time() - t0
        log(f"      [{key}] {u['기업명']}({u['소스']}): {len(tops)}건 근거 완료 "
            f"[{n}/{len(todo)}] {el/60:.1f}분 경과"
            + (f" · 남은 예상 {el/done*(need-done)/60:.0f}분" if done and not a.no_llm else ""))

    assert_no_promise(best, "보고서 JSON")
    if re.search(r"유망성\s*(점수|지수)|유망성\s*[:：]?\s*\d", json.dumps(best, ensure_ascii=False)):
        raise SystemExit("[중단] 보고서 JSON 본문에 유망성 점수 표현이 있습니다")
    if os.path.exists(a.out):
        os.replace(a.out, a.out + ".bak")
    json.dump(best, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    n_rec = sum(len(v["top10"]) for v in best.values())
    n_ex = sum(1 for v in best.values() for t in v["top10"] if t["추천근거_상세"])
    log(f"\n✔ {a.out}: 기업 {len(best)}건 · 추천 {n_rec}건 · 상세근거 {n_ex}건 "
        f"· 중복제외 과제 {len({t['과제고유번호'] for v in best.values() for t in v['top10']})}개")

    # 매칭 기준 비노출 점검: 생성분(재시도 후 잔존) / 최종 산출(치환 후)
    raw_v = sum(1 for k, v in ex.items() if BANNED.search(v)) \
        + sum(1 for k, v in rs.items() if BANNED.search(v))
    out_v = [(k, t["rank"]) for k, v in best.items() for t in v["top10"]
             if BANNED.search(t["추천근거_상세"]) or BANNED.search(t["판단근거"])]
    log(f"금지표현(매칭 기준) — 모델 출력 잔존 {raw_v}건(치환 처리) · 최종 산출 {len(out_v)}건"
        + (f" {out_v[:5]}" if out_v else ""))
    log(f"모델: {cm.MODEL_ID} · 계측 {STATS}")
    if n_ex < n_rec:
        log(f"! 상세근거 누락 {n_rec - n_ex}건 — 다시 실행하면 누락분만 생성한다")


if __name__ == "__main__":
    main()
