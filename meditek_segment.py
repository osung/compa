# -*- coding: utf-8 -*-
"""MEDITEK '주요사업분야(수요기술)' 원문을 선정 근거 구성요소로 분리.

원문이 [기술수요 | 회사소개 | 업종]을 섞어 담고 있어, 매칭 프롬프트에 무엇을
넣을지 선택할 수 있도록 문장 단위로 분류한다. 분류만 LLM(35B)이 하고 텍스트는
원문 조각을 그대로 보존한다(요약·생성 금지 — 근거 왜곡 방지).

  기술수요 : 확보·도입·공동개발하려는 기술, 해결 과제, 협력 희망 대상
  회사소개 : 당사가 무엇을 하는 기업인지에 대한 서술, 보유 제품·실적
  업종     : 업태/종목/산업분류 나열(기술 내용 없음)

사용: from meditek_segment import segment_text, SCOPES, compose
"""
import json
import re

import compa_match as cm

LABELS = ("기술수요", "회사소개", "업종")

# --demand-scope 값 → 포함할 라벨
SCOPES = {
    "tech": ("기술수요",),
    "tech+intro": ("기술수요", "회사소개"),
    "all": LABELS,
}

_SYS = ("당신은 기업이 제출한 '주요사업분야(수요기술)' 원문을 분석하는 기술이전 전문가입니다. "
        "원문의 각 문장을 성격에 따라 분류합니다. 문장을 요약하거나 새로 쓰지 않습니다.")

_GUIDE = (
    "위 원문의 각 문장을 아래 셋 중 하나로 분류하세요.\n"
    "  기술수요: 확보·도입·공동개발·이전받고자 하는 기술, 해결하려는 과제, 협력 희망 대상\n"
    "  회사소개: 당사가 어떤 기업인지에 대한 서술, 이미 보유한 제품·기술·실적\n"
    "  업종    : 업태·종목·산업분류의 단순 나열(구체적 기술 내용 없음)\n"
    "판단 기준: '무엇을 필요로 하는가'는 기술수요, '우리는 무엇을 하는가'는 회사소개, "
    "'어느 업종인가'는 업종입니다. 확신이 없으면 회사소개로 분류하세요.\n"
    "반드시 아래 JSON 형식으로만, 모든 문장에 대해 답하세요:\n"
    '{"results": [{"id": <문장번호>, "label": "<기술수요|회사소개|업종>"}, ...]}'
)


def split_sentences(text):
    """원문을 분류 단위(줄/문장)로 분리. 줄바꿈 우선, 긴 줄은 문장부호로 재분할."""
    units = []
    for ln in str(text or "").split("\n"):
        ln = ln.strip()
        if not ln:
            continue
        if len(ln) <= 120:
            units.append(ln)
            continue
        buf = ""
        for part in re.split(r"(?<=[.。!?])\s+|(?<=다)\s+(?=[-•\[])", ln):
            part = part.strip()
            if not part:
                continue
            if len(buf) + len(part) <= 120:
                buf = f"{buf} {part}".strip()
            else:
                if buf:
                    units.append(buf)
                buf = part
        if buf:
            units.append(buf)
    return units


def _parse_labels(out, n):
    m = re.search(r'\{.*"results".*\}', out, re.DOTALL)
    blob = m.group(0) if m else out
    res = {}
    try:
        for r in json.loads(blob).get("results", []):
            i, lab = int(r["id"]), str(r["label"]).strip()
            if 1 <= i <= n and lab in LABELS:
                res[i] = lab
    except Exception:
        for i, lab in re.findall(r'"id"\s*:\s*(\d+)\s*,\s*"label"\s*:\s*"([^"]+)"', blob):
            i, lab = int(i), lab.strip()
            if 1 <= i <= n and lab in LABELS:
                res[i] = lab
    return res


def segment_all_tech(text):
    """MEDITEK 기본 정책: '주요사업분야(수요기술)' 원문 전체를 기술수요로 간주.

    이 컬럼은 기업이 제출한 수요기술 서술이므로 회사소개/업종으로 떼어내지 않는다.
    (LLM 분류는 섹션 제목을 문장으로 오인하고 제목-본문 관계를 잃어 수요를 훼손했다.)
    """
    seg = {lab: [] for lab in LABELS}
    seg["기술수요"] = split_sentences(text)
    return seg


def segment_text(text):
    """원문 → {라벨: [문장, ...]}. LLM 실패분은 '회사소개'로 보수적 분류.

    ※ MEDITEK 원본에는 사용하지 않는다(segment_all_tech 사용). 다른 데이터셋에서
       회사소개·업종이 실제로 섞여 있을 때를 위해 남겨둔 경로.
    """
    units = split_sentences(text)
    if not units:
        return {lab: [] for lab in LABELS}
    block = "[원문 문장 목록]\n" + "\n".join(f"{i}. {u}" for i, u in enumerate(units, 1))
    msgs = [{"role": "system", "content": _SYS},
            {"role": "user", "content": f"{block}\n\n{_GUIDE}"}]
    try:
        labels = _parse_labels(
            cm.stream_explanation(msgs, max_tokens=1200, temperature=0.0, top_p=1.0),
            len(units))
    except Exception as e:
        print(f"      ! 세그먼트 분류 LLM 실패: {e}")
        labels = {}
    seg = {lab: [] for lab in LABELS}
    for i, u in enumerate(units, 1):
        seg[labels.get(i, "회사소개")].append(u)
    return seg


def compose(seg, scope):
    """선택한 라벨의 원문 조각만 이어붙여 매칭용 수요 텍스트 생성."""
    if scope not in SCOPES:
        raise ValueError(f"scope must be one of {sorted(SCOPES)}")
    parts = []
    for lab in SCOPES[scope]:
        parts.extend(seg.get(lab, []))
    return "\n".join(parts).strip()
