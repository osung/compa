# -*- coding: utf-8 -*-
"""
prompt_pipeline_demo.py

docs/prompt_keyword_extraction.md 와 docs/prompt_matching_score.md 의 스펙을 그대로
읽어들여, 사용자가 입력하는 요청을 처리하는 예제 파이프라인.

- 프롬프트(System / 가이드), 통합 불용어·연결어 목록은 **md 파일에서 직접 파싱**한다.
- 프롬프트 조립 → LLM 호출 → JSON 파싱 → (키워드) 규칙기반 폴백 / (점수) 배치·정렬은
  md 스펙 그대로 구현한다.
- 실제 Qwen 모델 호출부만 결정적 mock 으로 대체했다(환경변수 DEMO_REAL_LLM=1 이면
  compa_match.stream_explanation 로 교체 시도). mock 이라 하드웨어 없이 어디서나 돌아간다.

실행:  python3 examples/prompt_pipeline_demo.py
"""

import json
import os
import re
from pathlib import Path

KW_MIN, KW_MAX = 10, 20          # md '모델/파라미터'와 동일
DESC_CAND = 280                  # 후보 텍스트 절단 길이

DOCS = Path(__file__).resolve().parents[1] / "docs"
KW_MD = DOCS / "prompt_keyword_extraction.md"
SCORE_MD = DOCS / "prompt_matching_score.md"


# =====================================================================
# [0] md 스펙 로더 — 헤더 다음의 첫 ``` 코드블록을 뽑는다
# =====================================================================
def _code_block_after(md, needle):
    idx = md.find(needle)
    if idx < 0:
        raise ValueError(f"md에서 '{needle}' 헤더를 찾지 못함")
    fence = md.index("```", idx)
    start = md.index("\n", fence) + 1
    end = md.index("```", start)
    return md[start:end].rstrip("\n")


def load_specs():
    kw = KW_MD.read_text(encoding="utf-8")
    sc = SCORE_MD.read_text(encoding="utf-8")

    kw_sys = _code_block_after(kw, "## System prompt (`_KW_DOC_SYS`)")
    kw_guide = _code_block_after(kw, "### (b) 가이드")
    connectives = _code_block_after(kw, "### 연결어미·접속어")

    # 통합 불용어: 가이드 블록 안의 '아래 불용어 …' 줄부터 '반드시 아래 …' 줄 전까지
    g_lines = kw_guide.splitlines()
    s = next(i for i, l in enumerate(g_lines) if l.startswith("아래 불용어"))
    e = next(i for i, l in enumerate(g_lines) if l.startswith("반드시 아래"))
    stopwords = set()
    for l in g_lines[s + 1:e]:
        stopwords.update(w.strip() for w in l.split(",") if w.strip())

    conn = set()
    for l in connectives.splitlines():
        conn.update(w.strip() for w in l.split(",") if w.strip())

    sc_sys = _code_block_after(sc, "## System prompt (`_SYS_DOC`)")
    sc_guide = _code_block_after(sc, "### (c) 가이드")

    return {
        "kw_sys": kw_sys, "kw_guide": kw_guide,
        "stopwords": stopwords, "connectives": conn,
        "sc_sys": sc_sys, "sc_guide": sc_guide,
    }


SPEC = load_specs()


# =====================================================================
# [1] 규칙기반 명사 추출 (md '폴백' 처리 절차 1~7 그대로)
#     불용어·연결어는 md에서 파싱, 접미사/조사/동사어미는 md 프로즈 그대로 상수화
# =====================================================================
SUFFIXES = ("판매업", "도매업", "소매업", "서비스업", "제조업", "생산업",
            "가공업", "임대업", "대행업", "무역업", "유통업", "공업", "업")
JOSA = ("으로서", "으로써", "에서의", "에게서", "에서", "에게", "으로",
        "와의", "과의", "에의", "에", "의", "을", "를")
VERB_EOMI = ("하여", "하는", "되어", "되는", "한", "해", "된", "되", "함", "됨")


def is_meaningless_code(tok):
    """과기표준분류 코드('EH030301')·순수 숫자 등 의미 없는 ASCII 토큰 판정."""
    if re.search(r'[가-힣]', tok):
        return False
    if tok.isdigit():
        return True
    if len(tok) >= 5 and re.fullmatch(r'[A-Za-z0-9]+', tok) \
            and re.search(r'[A-Za-z]', tok) and re.search(r'\d', tok):
        return True
    return False


def normalize_text(text):
    """자유 문장 → 핵심 명사 키워드 리스트(등장 순서 유지). md 폴백 절차 1~7."""
    if not text:
        return []
    stop, conn = SPEC["stopwords"], SPEC["connectives"]
    text = re.sub(r'[^\w\s가-힣a-zA-Z0-9]', ' ', str(text))
    out = []
    for w in text.split():                                   # 1) 토큰화
        wc = w.casefold()
        if len(w) < 2 or wc in stop or wc in conn:           # 2) 1차 필터
            continue
        if is_meaningless_code(w):                           # 3) 무의미 코드
            continue
        stem = w
        for suf in SUFFIXES:                                 # 4) 업종 접미사
            if stem.endswith(suf) and len(stem) > len(suf):
                stem = stem[:-len(suf)]
                break
        for j in JOSA:                                       # 5) 어절 끝 조사
            if stem.endswith(j) and len(stem) - len(j) >= 2:
                stem = stem[:-len(j)]
                break
        drop = False                                         # 6) 불용어 동사 활용형
        for em in VERB_EOMI:
            if stem.endswith(em) and stem[:-len(em)].casefold() in stop:
                drop = True
                break
        if drop:
            continue
        stem = stem.strip()
        sc = stem.casefold()
        if len(stem) < 2 or sc in stop or sc in conn:        # 7) 재검사
            continue
        out.append(stem)
    return out


# =====================================================================
# [2] LLM 호출부 — 기본 mock(결정적). DEMO_REAL_LLM=1 이면 실제 백엔드 시도.
# =====================================================================
def _real_llm(messages):
    from compa_match import stream_explanation
    return stream_explanation(messages, max_tokens=3500, temperature=0.0, top_p=1.0)


def _mock_llm(messages):
    """system 내용으로 작업을 구분해 그럴듯한 JSON을 결정적으로 생성."""
    sys = messages[0]["content"]
    user = messages[1]["content"]
    if "핵심 기술 키워드를 추출" in sys:
        return _mock_keywords(user)
    return _mock_scores(user)


def call_llm(messages):
    if os.environ.get("DEMO_REAL_LLM") == "1":
        return _real_llm(messages)
    return _mock_llm(messages)


# 가이드가 '키워드로 삼지 말라'고 명시한 청구항/행정 상용문구 (LLM·폴백 공통 최종 배제)
BOILERPLATE = {"특징", "상기", "것을", "것", "포함하는", "그것"}
_TAIL = ("으로", "는", "은", "와", "과", "이", "가", "을", "를", "의", "도", "로")


def _clean_tail(tok):
    """조사 꼬리를 정리해 깨끗한 명사형으로(예: 음극재는→음극재, 구조와→구조).
    LLM이 명사 키워드를 출력하는 동작을 mock 에서 흉내내기 위함."""
    if not re.search(r'[가-힣]', tok):
        return tok
    for t in _TAIL:
        if tok.endswith(t) and len(tok) - len(t) >= 2:
            return tok[:-len(t)]
    return tok


def _mock_keywords(user):
    """가짜 LLM(가이드 준수): 문서의 핵심 명사를 정리해 '중요도(빈도·길이) 순'으로 반환.
    청구항 상용문구는 가이드대로 배제한다."""
    doc = user.split("\n\n위 [문서]")[0].replace("[문서]\n", "", 1)
    toks = [_clean_tail(t) for t in normalize_text(doc)]
    freq = {}
    for t in toks:
        freq[t] = freq.get(t, 0) + 1
    picked, seen = [], set()
    for t in sorted(dict.fromkeys(toks), key=lambda x: (-freq[x], -len(x))):
        c = t.casefold()
        if len(t) < 2 or c in seen or c in SPEC["stopwords"] or t in BOILERPLATE:
            continue
        seen.add(c)
        picked.append(t)
    return json.dumps({"keywords": picked[:KW_MAX]}, ensure_ascii=False)


def _mock_scores(user):
    """가짜 LLM: 기준 문서와 후보 문서의 명사 겹침(Dice)으로 0~100 점수."""
    src_text = user.split("\n\n[평가 대상")[0]
    src_text = re.sub(r'^\[[^\]]*\]\n', '', src_text)
    src = set(w.casefold() for w in normalize_text(src_text))
    cand_block = user.split("[평가 대상", 1)[1]
    results = []
    for m in re.finditer(r'^(\d+)\.\s*(.*)$', cand_block, re.MULTILINE):
        cid, ctext = int(m.group(1)), m.group(2)
        ck = set(w.casefold() for w in normalize_text(ctext))
        inter = src & ck
        denom = (len(src) + len(ck)) or 1
        score = int(round(min(100, (2 * len(inter) / denom) * 145)))
        if inter:
            shown = [w for w in normalize_text(ctext) if w.casefold() in inter][:3]
            reason = "핵심 기술 부합: " + ", ".join(shown)
        else:
            reason = "직접 관련 근거 약함"
        results.append({"id": cid, "score": max(0, min(100, score)),
                        "reason": reason[:40]})
    return json.dumps({"results": results}, ensure_ascii=False)


# =====================================================================
# [3] 키워드 추출 (md 프롬프트 조립 + 파싱 + 폴백)
# =====================================================================
def _parse_keywords(text):
    m = re.search(r'\{.*"keywords".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("keywords", []) if isinstance(data, dict) else data
    except Exception:
        mm = re.search(r'"keywords"\s*:\s*\[(.*?)\]', text, re.DOTALL)
        items = re.findall(r'"([^"]+)"', mm.group(1)) if mm else []
    seen, out = set(), []
    for it in items:
        v = str(it).strip().strip('"').strip()
        if v and v.casefold() not in seen:
            seen.add(v.casefold())
            out.append(v)
    return out[:KW_MAX]


def extract_keywords_doc(text):
    """임의의 기술사업화 문서 문자열 → 중요도 순 키워드 리스트."""
    text = str(text or "").strip()
    user = f"[문서]\n{text}\n\n{SPEC['kw_guide']}"
    msgs = [{"role": "system", "content": SPEC["kw_sys"]},
            {"role": "user", "content": user}]
    try:
        kws = _parse_keywords(call_llm(msgs))
    except Exception as e:
        print(f"   ! 키워드 LLM 실패: {e}")
        kws = []
    if len(kws) < KW_MIN:                     # 규칙기반 폴백(뒤에 이어붙임)
        seen = {k.casefold() for k in kws}
        for v in (_clean_tail(t) for t in normalize_text(text)):
            if v.casefold() not in seen:
                seen.add(v.casefold())
                kws.append(v)
            if len(kws) >= KW_MAX:
                break
    # 가이드의 청구항/행정 상용문구 배제를 두 경로(LLM·폴백) 최종 산출에 일괄 적용
    kws = [k for k in kws if k not in BOILERPLATE]
    return kws[:KW_MAX]


# =====================================================================
# [4] 적합도 점수 (md 프롬프트 조립 + 배치 + 파싱 + 정렬)
# =====================================================================
LLM_BATCH = 20


def _parse_scores(text):
    out = {}
    m = re.search(r'\{.*"results".*\}', text, re.DOTALL)
    blob = m.group(0) if m else text
    try:
        data = json.loads(blob)
        items = data.get("results", []) if isinstance(data, dict) else data
    except Exception:                          # 개별 객체 정규식 폴백
        items = [{"id": int(x.group(1)), "score": int(x.group(2)), "reason": x.group(3)}
                 for x in re.finditer(
                     r'\{\s*"id"\s*:\s*(\d+)\s*,\s*"score"\s*:\s*(\d+)\s*,\s*"reason"\s*:\s*"([^"]*)"',
                     text)]
    for it in items:
        try:
            i = int(it["id"])
            s = max(0, min(100, int(round(float(it["score"])))))
            out[i] = (s, str(it.get("reason", "")).strip()[:120])
        except (KeyError, ValueError, TypeError):
            continue
    return out


def score_candidates(source, cands, src_label, cand_label, cand_maxlen=DESC_CAND):
    """source(문자열) 대비 후보 리스트[{id,text}]를 0~100 평가 → 점수 내림차순 정렬."""
    guide = SPEC["sc_guide"].replace("{src_label}", src_label).replace("{cand_label}", cand_label)
    scores = {}
    for b in range(0, len(cands), LLM_BATCH):
        batch = cands[b:b + LLM_BATCH]
        cand_lines = [f"[평가 대상 {cand_label} 목록]"]
        for c in batch:
            cand_lines.append(f"{c['id']}. {str(c['text']).strip()[:cand_maxlen]}")
        user = f"[{src_label}]\n{source}\n\n" + "\n".join(cand_lines) + f"\n\n{guide}"
        msgs = [{"role": "system", "content": SPEC["sc_sys"]},
                {"role": "user", "content": user}]
        try:
            parsed = _parse_scores(call_llm(msgs))
        except Exception as e:
            print(f"   ! 점수 LLM 실패: {e}")
            parsed = {}
        for c in batch:
            scores[c["id"]] = parsed.get(c["id"], (0, "(LLM 응답 누락)"))
    ranked = sorted(cands, key=lambda c: scores[c["id"]][0], reverse=True)
    return [(c, scores[c["id"]][0], scores[c["id"]][1]) for c in ranked]


# =====================================================================
# [5] 가짜 데이터 (그럴듯하게)
# =====================================================================
COMPANIES = {
    "C1": "나노셀텍은 리튬이온전지용 실리콘 복합 음극재를 개발하는 소재 기업이다. "
          "고용량·장수명 배터리를 위한 실리콘-탄소 나노복합 구조와 전해질 첨가제 기술을 보유한다.",
    "C2": "바이오팜코리아는 도라지 유래 사포닌 Platycodin D의 생체이용률(Bioavailability)을 "
          "높이는 나노에멀젼 제형 기술을 보유한 천연물 의약품 기업이다.",
    "C3": "그린모빌리티는 수소연료전지용 백금 촉매와 고분자 전해질막(MEA)을 제조하는 회사다.",
    "C4": "루미넥스디스플레이는 차세대 OLED 청색 인광 발광소재와 정공수송 소재를 합성하는 "
          "디스플레이 소재 기업이다.",
    "C5": "젠에딧은 CRISPR 유전자편집 기반 작물 형질전환 및 세포주 개발 플랫폼을 보유한 바이오 기업이다.",
}
PROJECTS = {
    "P1": "리튬이차전지 고용량 실리콘 음극재 및 전해질 첨가제 상용화 기술 개발",
    "P2": "천연물 사포닌 Platycodin D 생체이용률 개선 나노제형 연구",
    "P3": "수소연료전지 백금 촉매 및 전해질막 내구성 향상 기술 개발",
    "P4": "차세대 OLED 청색 인광 발광소재 합성 기술 개발",
    "P5": "CRISPR 유전자편집 기반 작물 형질전환 플랫폼 구축",
}
PATENTS = {
    "PT1": "실리콘-탄소 복합 음극재 및 그 제조방법. 상기 음극재는 탄소 매트릭스에 분산된 "
           "실리콘 나노입자를 포함하는 것을 특징으로 하는 리튬이차전지용 음극재.",
    "PT2": "Platycodin D를 포함하는 나노에멀젼 조성물로서, 생체이용률을 향상시키는 것을 "
           "특징으로 하는 경구용 제제.",
    "PT3": "수소저장용 마그네슘계 합금 및 상기 합금을 이용하는 수소 저장 장치.",
    "PT4": "청색 인광 OLED 발광소재 화합물로서, 정공수송 특성을 갖는 것을 특징으로 하는 "
           "유기발광 화합물.",
    "PT5": "CRISPR-Cas9 기반 유전자편집 방법 및 상기 방법으로 형질전환된 식물체.",
}
DEMANDS = {
    "D1": "실리콘 음극재 고용량 배터리 수명 개선",           # 짧음 → 폴백 유도
    "D2": "수소연료전지 백금 촉매 내구성 향상 전해질막",     # 짧음 → 폴백 유도
}


def _cands(dct):
    """{key: text} → 데모용 후보 리스트 [{id, key, text}] (id=1..N)."""
    out = []
    for i, (k, v) in enumerate(dct.items(), 1):
        out.append({"id": i, "key": k, "text": v})
    return out


# =====================================================================
# [6] 사용자 요청 처리 (여러 시나리오)
# =====================================================================
def handle(req):
    print("=" * 78)
    print(f"[요청] {req['title']}")
    if req["task"] == "keywords":
        print(f"  · 문서유형: {req['doc_type']}")
        print(f"  · 원문: {req['text'][:70]}{'…' if len(req['text'])>70 else ''}")
        kws = extract_keywords_doc(req["text"])
        print(f"  → 키워드({len(kws)}개, 중요도 순): {', '.join(kws)}")
        return {"name": req["name"], "keywords": kws}
    else:
        cands = req["candidates"]
        print(f"  · 방향: [{req['src_label']}] → [{req['cand_label']}]  (후보 {len(cands)}건)")
        ranked = score_candidates(req["source"], cands,
                                  req["src_label"], req["cand_label"])
        print("  → 적합도 순위:")
        for rank, (c, sc, reason) in enumerate(ranked, 1):
            print(f"     {rank}. [{c['key']}] score={sc:3d}  {reason}")
        return {"name": req["name"], "ranked": ranked}


def build_requests():
    proj = _cands(PROJECTS)
    comp = _cands(COMPANIES)
    pat = _cands(PATENTS)
    return [
        # ---- 키워드 추출 시나리오 (유형별 2건씩) ----
        {"name": "kw_company_1", "task": "keywords", "title": "기업 소개문 키워드 추출 #1",
         "doc_type": "기업", "text": COMPANIES["C1"]},
        {"name": "kw_company_2", "task": "keywords", "title": "기업 소개문 키워드 추출 #2",
         "doc_type": "기업", "text": COMPANIES["C4"]},
        {"name": "kw_patent_1", "task": "keywords", "title": "특허 명세(청구항 상용문구 배제) 키워드 추출 #1",
         "doc_type": "특허", "text": PATENTS["PT1"]},
        {"name": "kw_patent_2", "task": "keywords", "title": "특허 명세(청구항 상용문구 배제) 키워드 추출 #2",
         "doc_type": "특허", "text": PATENTS["PT2"]},
        {"name": "kw_project_1", "task": "keywords", "title": "국가 R&D 과제명 키워드 추출 #1",
         "doc_type": "국가 R&D 과제", "text": PROJECTS["P2"]},
        {"name": "kw_project_2", "task": "keywords", "title": "국가 R&D 과제명 키워드 추출 #2",
         "doc_type": "국가 R&D 과제", "text": PROJECTS["P4"]},
        {"name": "kw_demand_1", "task": "keywords", "title": "짧은 수요기술 → 규칙기반 폴백 보충 #1",
         "doc_type": "수요기술", "text": DEMANDS["D1"]},
        {"name": "kw_demand_2", "task": "keywords", "title": "짧은 수요기술 → 규칙기반 폴백 보충 #2",
         "doc_type": "수요기술", "text": DEMANDS["D2"]},

        # ---- 적합도 점수(매칭) 시나리오 (방향별 2건씩) ----
        {"name": "m_comp2proj_1", "task": "match", "title": "기업 → 국가 R&D 과제 매칭 #1",
         "src_label": "기업", "source": COMPANIES["C1"],
         "cand_label": "국가 R&D 과제", "candidates": proj},
        {"name": "m_comp2proj_2", "task": "match", "title": "기업 → 국가 R&D 과제 매칭 #2",
         "src_label": "기업", "source": COMPANIES["C3"],
         "cand_label": "국가 R&D 과제", "candidates": proj},
        {"name": "m_proj2comp_1", "task": "match", "title": "국가 R&D 과제 → 기업 매칭(역방향) #1",
         "src_label": "국가 R&D 과제", "source": PROJECTS["P3"],
         "cand_label": "기업", "candidates": comp},
        {"name": "m_proj2comp_2", "task": "match", "title": "국가 R&D 과제 → 기업 매칭(역방향) #2",
         "src_label": "국가 R&D 과제", "source": PROJECTS["P4"],
         "cand_label": "기업", "candidates": comp},
        {"name": "m_comp2pat_1", "task": "match", "title": "기업 → 특허 매칭 #1",
         "src_label": "기업", "source": COMPANIES["C2"],
         "cand_label": "특허", "candidates": pat},
        {"name": "m_comp2pat_2", "task": "match", "title": "기업 → 특허 매칭 #2",
         "src_label": "기업", "source": COMPANIES["C1"],
         "cand_label": "특허", "candidates": pat},
        {"name": "m_demand2proj_1", "task": "match", "title": "수요기술 → 국가 R&D 과제 매칭 #1",
         "src_label": "수요기술", "source": DEMANDS["D1"],
         "cand_label": "국가 R&D 과제", "candidates": proj},
        {"name": "m_demand2proj_2", "task": "match", "title": "수요기술 → 국가 R&D 과제 매칭 #2",
         "src_label": "수요기술", "source": DEMANDS["D2"],
         "cand_label": "국가 R&D 과제", "candidates": proj},
    ]


# =====================================================================
# [7] 자동 점검 (정상 동작 확인용 assert)
# =====================================================================
def self_check(outputs):
    checks = []
    by = {o["name"]: o for o in outputs}          # 이름 기반 조회(인덱스 하드코딩 제거)

    def ok(name, cond):
        checks.append((name, bool(cond)))

    def kws(n):
        return by[n]["keywords"]

    def top(n):
        return by[n]["ranked"][0][0]["key"]

    def scores(n):
        return [s for _, s, _ in by[n]["ranked"]]

    # 공통: 모든 키워드 산출물의 개수 상한·중복
    for o in outputs:
        if "keywords" not in o:
            continue
        ok(f"[{o['name']}] 키워드 개수 ≤ KW_MAX", len(o["keywords"]) <= KW_MAX)
        ok(f"[{o['name']}] 키워드 중복 없음",
           len(o["keywords"]) == len({k.casefold() for k in o["keywords"]}))

    # 공통: 모든 매칭 산출물의 점수 범위·정렬·후보 누락
    for o in outputs:
        if "ranked" not in o:
            continue
        scs = [s for _, s, _ in o["ranked"]]
        ok(f"[{o['name']}] 점수 0~100 범위", all(0 <= s <= 100 for s in scs))
        ok(f"[{o['name']}] 점수 내림차순 정렬", scs == sorted(scs, reverse=True))

    # 특허 키워드 2건: 청구항 상용문구 배제
    for n in ("kw_patent_1", "kw_patent_2"):
        ok(f"[{n}] '상기'·'특징' 배제", all("상기" not in k and "특징" not in k for k in kws(n)))
    ok("[kw_patent_1] 핵심어 포함(실리콘/음극재 등)",
       any(k in kws("kw_patent_1") for k in ("실리콘", "음극재", "나노입자", "탄소")))
    ok("[kw_patent_2] 핵심어 포함(Platycodin/나노에멀젼 등)",
       any(k in kws("kw_patent_2") for k in ("Platycodin", "나노에멀젼", "생체이용률", "경구용")))

    # 기업 키워드 2건: 핵심어 포함
    ok("[kw_company_1] 핵심어 포함", any(k in kws("kw_company_1") for k in ("실리콘", "음극재", "배터리")))
    ok("[kw_company_2] 핵심어 포함", any(k in kws("kw_company_2") for k in ("OLED", "인광", "발광소재", "청색")))

    # 과제 키워드 2건: 핵심어 포함
    ok("[kw_project_1] 핵심어 포함", any(k in kws("kw_project_1") for k in ("Platycodin", "생체이용률", "나노제형", "사포닌")))
    ok("[kw_project_2] 핵심어 포함", any(k in kws("kw_project_2") for k in ("OLED", "인광", "발광소재", "청색")))

    # 짧은 수요기술 2건: 폴백으로 최소 확보
    for n in ("kw_demand_1", "kw_demand_2"):
        ok(f"[{n}] 짧은 문서도 키워드 ≥ 3", len(kws(n)) >= 3)

    # 매칭 순위 정확성 (방향별 2건씩)
    ok("[m_comp2proj_1] C1(실리콘 음극재) → P1 1위", top("m_comp2proj_1") == "P1")
    ok("[m_comp2proj_2] C3(수소연료전지) → P3 1위", top("m_comp2proj_2") == "P3")
    ok("[m_proj2comp_1] P3(수소연료전지) → C3 1위", top("m_proj2comp_1") == "C3")
    ok("[m_proj2comp_2] P4(OLED) → C4 1위", top("m_proj2comp_2") == "C4")
    ok("[m_comp2pat_1] C2(Platycodin) → PT2 1위", top("m_comp2pat_1") == "PT2")
    ok("[m_comp2pat_2] C1(실리콘) → PT1 1위", top("m_comp2pat_2") == "PT1")
    ok("[m_demand2proj_1] D1(실리콘) → P1 1위", top("m_demand2proj_1") == "P1")
    ok("[m_demand2proj_2] D2(수소) → P3 1위", top("m_demand2proj_2") == "P3")

    # 후보 누락 없음(점수 dict가 모든 후보를 커버)
    ok("[m_comp2proj_1] 후보 5건 모두 순위화", len(by["m_comp2proj_1"]["ranked"]) == 5)

    # 파서 견고성: 깨진 JSON도 정규식 폴백으로 복구
    broken = '어쩌구 {"results": [{"id": 1, "score": 88, "reason": "ok"} 잘림'
    ok("점수 파서 정규식 폴백", _parse_scores(broken).get(1) == (88, "ok"))

    print("\n" + "#" * 78)
    print("자동 점검 결과")
    print("#" * 78)
    allok = True
    for name, passed in checks:
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")
        allok = allok and passed
    print("-" * 78)
    print("결과:", "✅ ALL CHECKS PASSED" if allok else "❌ SOME CHECKS FAILED")
    return allok


def main():
    backend = "REAL(compa_match)" if os.environ.get("DEMO_REAL_LLM") == "1" else "MOCK(결정적)"
    print(f"LLM 백엔드: {backend}")
    print(f"로드한 스펙: 불용어 {len(SPEC['stopwords'])}개 / 연결어 {len(SPEC['connectives'])}개\n")
    outputs = [handle(req) for req in build_requests()]
    allok = self_check(outputs)
    raise SystemExit(0 if allok else 1)


if __name__ == "__main__":
    main()
