# -*- coding: utf-8 -*-
"""보고서에 싣는 텍스트의 표기 오류(띄어쓰기·오탈자) 교정.

기업이 제출한 hwp 원문에는 줄바꿈 자리가 공백으로 굳어 단어가 쪼개진 흔적이 남아 있다
('향상시킵 니다', '직 접 희석', '이 미지 분석'). 보고서에 그대로 실리면 눈에 띄므로
렌더링 전에 교정한다. 세 층으로 나누고, 뒤로 갈수록 판단이 필요한 일을 맡는다.

  ① marks  : 제어문자·중복공백·구두점·괄호·물결표·숫자와 단위 — 결정적 규칙(오탐 없음)
  ② table  : 실제로 확인한 오탈자·긴 단어 분리 사례표(맥락까지 확인한 것만 넣는다)
  ③ join   : 형태소 분석(Kiwi)으로 '단어 안에 들어간 공백'만 골라 없앤다

③ 이 이 모듈의 핵심이다. Kiwi 의 space() 로 문장을 다시 띄어 쓰는 방식은 이 도메인에서
쓸 수 없었다 — '스크리닝'을 '스 크 리닝'으로 쪼개고 '의료기기'를 '의료 기기'로 벌리는 등
고치는 것보다 망가뜨리는 것이 많았다. 그래서 공백을 넣는 일은 하지 않고 **없애는 일만**
하며, 아래 세 조건을 모두 만족할 때만 없앤다.

  (가) 공백을 없앤 문장의 분석 점수가 원문보다 JOIN_MIN 이상 높다
  (나) 없앤 자리를 한 형태소가 가로지른다(= 공백이 단어 안에 있었다)
       그 형태소가 4음절 이상이면 저자가 띄어 쓴 복합어로 보고 건드리지 않는다
       ('엔드 이펙터'·'유도 가열'·'형광 이미징' 은 정상 표기다)
  (다) 또는 오른쪽 조각이 조사·어미·접미사뿐이다(낱말 앞에 올 수 없는 형태소)

이 규칙을 260820 매칭 보고서 전체(26개사·130건, 순한글 공백 4만4천 곳)에 걸어 후보 50건을
모두 눈으로 확인했다 — 50건 전부 올바른 교정이고, 정상 복합어를 붙인 사례는 없었다.
Kiwi 가 없으면 ③ 만 건너뛰고 ①②는 그대로 동작한다.

사용:
  import report_text_fix as tf
  tf.fix("신약 스크리닝 효율을 혁신적으로 향상시킵 니다")   # → '…향상시킵니다'
  tf.fix_report(demands, patents)        # 보고서 JSON·특허 dict 를 제자리 교정
"""
import re

try:                                        # 형태소 분석기는 선택 의존(없으면 ③ 생략)
    from kiwipiepy import Kiwi
    _KIWI = Kiwi()
except Exception as _e:                      # pragma: no cover
    _KIWI, _KIWI_ERR = None, _e

JOIN_MIN = 5.0      # 공백 제거로 얻어야 하는 최소 점수 개선(4.8 이하에는 오탐이 섞였다)
MAX_SPAN = 3        # 공백을 품은 형태소가 이보다 길면 정상 복합어로 본다(음절)
# 오른쪽 조각이 한 음절이고 아래에 들어가면 붙이지 않는다 — 관형사·부사로도 쓰여 맥락이
# 갈린다('… 등 이 기업의' 을 '등이 기업의' 로 바꾸면 뜻이 달라진다).
AMBIGUOUS_RIGHT = {"이", "그", "저", "및", "등", "또", "안", "못", "잘", "더", "한"}
GLUE_TAGS = ("J", "E")                       # 조사·어미
GLUE_PREFIX = ("XS",)                        # 접미사(-적/-화/-용/-하다 …)

STATS = {"marks": 0, "table": 0, "join": 0}   # 교정 건수 계측(검증 로그용)

# ---- ① 결정적 규칙 ----------------------------------------------------------
_CTRL = re.compile(r"_x00[0-9A-Fa-f]{2}_|[\x00-\x08\x0b\x0c\x0e-\x1f]")
_ZW = re.compile(r"[​-‍﻿]")
# 숫자 뒤 단위·의존명사 — 한국어 조판에서 붙여 쓰는 것만 넣는다.
# Hz·nm·N 같은 라틴 SI 단위는 숫자와 띄어 쓰는 것이 표준이므로 건드리지 않는다.
_NUM_UNIT = re.compile(r"(\d)\s+(년|개월|주|일|시간|분|초|건|명|개|종|배|차|회|위|원|억|만|"
                       r"차원|%)(?=[가-힣%)\]]|$|\s)")
# hwp 잔재: '3 D' → '3D'
_NUM_D = re.compile(r"(\d)\s+(D)(?![A-Za-z가-힣])")
_TILDE = re.compile(r"(\d)\s*~\s*(\d)")       # '40 ~4000' → '40~4000'
_PAREN_L, _PAREN_R = re.compile(r"\(\s+"), re.compile(r"\s+\)")
_SP_PUNCT = re.compile(r"[ \t]+([,.;:!?)\]])")    # 구두점 앞 공백
# 닫는 괄호·라틴문자·숫자 뒤에 떨어진 조사('(40~4000nL) 와' → '(40~4000nL)와',
# 'MiRI 는' → 'MiRI는', 'AI 의' → 'AI의'). 조사 뒤가 공백·구두점·문장끝일 때만 붙여
# '부속(가) 도구를' 처럼 낱말로 시작하는 경우를 건드리지 않는다.
# '이' 는 넣지 않는다 — '[기술 적합성] 이 과제는' 의 '이' 는 조사가 아니라 관형사라서
# 붙이면 '[기술 적합성]이 과제는' 이 되어 뜻이 망가진다.
_TAIL_JOSA = re.compile(r"([)\]”’A-Za-z0-9%])[ \t]+"
                        r"((?:으로|에서|에게|부터|까지|[은는가을를와과의로도])(?=[\s,.;:!?]|$))")
_DUP_PUNCT = re.compile(r"([,;:])(\s*\1)+")       # ', ,' → ','
_DUP_SP = re.compile(r"[ \t]{2,}")
# 어미가 떨어진 자리 — 한국어에 '습니다'·'니다' 로 시작하는 낱말은 없으므로 점수 판정 없이
# 붙여도 안전하다('향상시킵 니다' / '해결하였 습니다'). 뒤가 공백·구두점·문장끝일 때만.
_ENDING = re.compile(r"(?<=[가-힣])[ \t]+(습니다|습니까|니다|니까)(?=[\s.,;:!?)\]”’]|$)")
# '검증되었습니 다' 처럼 어미 안쪽이 갈라진 경우. '습니/입니' 뒤로 한정해 '그러니 다시'
# 같은 정상 표기를 건드리지 않는다.
_ENDING_IN = re.compile(r"(?<=습니)[ \t]+([다까])(?=[\s.,;:!?)\]”’]|$)"
                        r"|(?<=입니)[ \t]+([다까])(?=[\s.,;:!?)\]”’]|$)")
_PERIOD = re.compile(r"([가-힣%)\]』」’”])\.(?=[^\s.,;:)\]}%’”」』])")   # '…다.이' → '…다. 이'
_PUBLIC_RND = re.compile(r"공공\s*R&D")


def _marks(s):
    out = _ZW.sub("", _CTRL.sub("", s))
    out = out.replace("\xa0", " ")
    out = _PUBLIC_RND.sub(lambda m: m.group(0).replace("공공", "국가"), out)
    out = _PAREN_L.sub("(", _PAREN_R.sub(")", out))
    out = _NUM_D.sub(r"\1\2", _NUM_UNIT.sub(r"\1\2", out))
    out = _TILDE.sub(r"\1~\2", out)
    out = _DUP_PUNCT.sub(r"\1", _SP_PUNCT.sub(r"\1", out))
    out = _TAIL_JOSA.sub(r"\1\2", out)
    out = _DUP_SP.sub(" ", out)
    out = _ENDING.sub(r"\1", out)
    out = _ENDING_IN.sub(lambda m: m.group(1) or m.group(2), out)
    out = _PERIOD.sub(r"\1. ", out)
    return "\n".join(ln.rstrip() for ln in out.split("\n"))


# ---- ② 확인한 사례표 --------------------------------------------------------
# 원문을 직접 열어 맥락까지 확인한 것만 넣는다. 자동 추론은 하지 않는다 —
# '국재화(局在化)' 처럼 사전에 약한 정상 용어를 오타로 바꿔 버리는 사고를 막기 위한 방침.
CORRECTIONS = {
    # 오탈자 (제출 원문)
    "간겅가능식품": "건강기능식품",      # '수면관리 간겅가능식품' — 같은 문장의 앞부분이 '건강기능식품'
    "감연 관리": "감염 관리",            # '세척 팁의 일회용화로 감연 관리 효율성'
    # 4음절 이상이라 ③ 규칙이 정상 복합어와 구분하지 못하는 단어 분리(hwp 줄바꿈 잔재)
    "알고리 즘": "알고리즘",
    "포트폴 리오": "포트폴리오",
    "오가 노이드": "오가노이드",
    # 조사 오용 — '배열'은 받침이 있어 '과'가 맞다. 일반 규칙으로 처리하지 않는다:
    # '효과'·'치과'·'전계효과' 처럼 '과'가 낱말 안에 든 경우와 규칙만으로는 구분되지 않는다.
    "크리스탈 배열 와": "크리스탈 배열과",
    # 35B 가 근거문에서 기관명을 잘못 옮긴 경우. 원본 데이터(과제수행기관·공급기관·
    # 과제설명문)는 모두 '숭실대학교'로 정확하다 — 생성 단계의 오류다.
    # verify_report_supply 의 '기관명' 검사가 같은 유형을 다시 잡는다.
    "술흘대학교": "숭실대학교",
}
_TABLE = re.compile("|".join(re.escape(k) for k in
                             sorted(CORRECTIONS, key=len, reverse=True))) if CORRECTIONS else None


def _table(s):
    return _TABLE.sub(lambda m: CORRECTIONS[m.group(0)], s) if _TABLE else s


# ---- ③ 형태소 기반 공백 제거 -------------------------------------------------
_SPACE = re.compile(r"(?<=[가-힣])([ \t])(?=[가-힣])")
_WORD = re.compile(r"\S+")
_sc, _tk = {}, {}


def _score(s):
    if s not in _sc:
        r = _KIWI.analyze(s, top_n=1)
        _sc[s] = r[0][1] if r else -1e9
    return _sc[s]


def _tokens(s):
    if s not in _tk:
        _tk[s] = _KIWI.tokenize(s)
    return _tk[s]


def _is_glue(tag):
    return tag[0] in GLUE_TAGS or tag.startswith(GLUE_PREFIX)


def _should_join(seg, j):
    """seg[j] 공백을 없앨지 판정 → (여부, 이유)."""
    cand = seg[:j] + seg[j + 1:]
    if _score(cand) - _score(seg) < JOIN_MIN:
        return False, "점수 개선 없음"
    ts = _tokens(cand)
    span = [t for t in ts if t.start < j < t.start + len(t.form)]
    if span:
        t = span[0]
        if len(t.form) > MAX_SPAN and not _is_glue(t.tag):
            return False, f"정상 복합어 {t.form}/{t.tag}"
        if t.tag == "EF":                    # 종결어미가 문장 중간에 생기면 오판
            tail = cand[t.start + len(t.form):].lstrip()
            if tail and tail[0] not in ".!?":
                return False, f"종결어미 {t.form} 뒤 본문 계속"
        return True, f"형태소 내부 {t.form}/{t.tag}"
    if not [t for t in ts if t.start == j]:
        return False, "경계 불명"
    m = _WORD.match(seg[j + 1:])
    right = m.group(0) if m else ""
    if len(right) == 1 and right in AMBIGUOUS_RIGHT:
        return False, f"모호한 1음절 '{right}'"
    # 오른쪽 조각을 따로 떼어 분석하면 안 된다 — '습니다' 만 넣으면 '슬/VV + ᆸ니다/EF' 로
    # 읽혀 어미가 아닌 것처럼 보인다. 붙인 문장의 분석에서 그 구간의 형태소를 본다.
    rts = [t for t in ts if j <= t.start < j + len(right)]
    if rts and all(_is_glue(t.tag) for t in rts):
        return True, "조사·어미 " + "+".join(f"{t.form}/{t.tag}" for t in rts)
    return False, "우측이 독립어"


def _join(s, log=None):
    """단어 안에 들어간 공백만 제거. 한 줄씩, 개선폭이 큰 자리부터 반복 적용."""
    if _KIWI is None:
        return s
    out = []
    for line in s.split("\n"):
        segs = re.split(r"(?<=[.!?])\s+", line) if len(line) > 400 else [line]
        fixed = []
        for seg in segs:
            for _ in range(8):               # 한 문장에 여러 곳이 있을 수 있다
                best = None
                for m in _SPACE.finditer(seg):
                    ok, why = _should_join(seg, m.start(1))
                    if not ok:
                        continue
                    d = _score(seg[:m.start(1)] + seg[m.end(1):]) - _score(seg)
                    if best is None or d > best[0]:
                        best = (d, m.start(1), why)
                if best is None:
                    break
                _, j, why = best
                if log is not None:
                    log.append((seg[max(0, j - 10):j] + "␣" + seg[j + 1:j + 11], why))
                seg = seg[:j] + seg[j + 1:]
                STATS["join"] += 1
            fixed.append(seg)
        out.append(" ".join(fixed) if len(segs) > 1 else fixed[0])
    return "\n".join(out)


# ---- 진입점 -----------------------------------------------------------------
def fix(text, log=None):
    """보고서에 실을 한 덩어리 텍스트를 교정한다(빈 값은 그대로)."""
    s = "" if text is None else str(text)
    if not s.strip():
        return s
    a = _marks(s)
    STATS["marks"] += (a != s)
    b = _table(a)
    STATS["table"] += (b != a)
    return _join(b, log)


# 보고서에 '보이는' 필드만 교정한다. 과제설명문은 gen_report 의 fmt_period/extract_class 가
# 정규식으로 파싱하는 입력이라 손대지 않는다(표시에는 쓰이지 않는다).
COMPANY_FIELDS = ("수요기술명", "수요기술 내용", "기보유기술명", "기보유기술 내용",
                  "기술유형", "기술분야")
TOP_FIELDS = ("과제명", "판단근거", "추천근거_상세", "수행기관", "공급기관")
PATENT_FIELDS = ("특허명", "기관")


def fix_report(demands, patents=None, top_key="top5", log=None):
    """보고서 입력(JSON dict)과 특허 dict 를 제자리 교정 → 교정 로그 리스트."""
    log = [] if log is None else log
    for v in demands.values():
        for k in COMPANY_FIELDS:
            if k in v:
                v[k] = fix(v[k], log)
        for t in v.get(top_key, []):
            for k in TOP_FIELDS:
                if k in t:
                    t[k] = fix(t[k], log)
    for lst in (patents or {}).values():
        for p in lst:
            for k in PATENT_FIELDS:
                if k in p:
                    p[k] = fix(p[k], log)
    return log


def report_stats():
    engine = "kiwipiepy" if _KIWI else f"형태소 분석기 없음({_KIWI_ERR})"
    return f"표기 교정: 기호·간격 {STATS['marks']}필드 · 사례표 {STATS['table']}필드 · " \
           f"단어 내 공백 {STATS['join']}곳 ({engine})"
