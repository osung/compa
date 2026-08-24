# -*- coding: utf-8 -*-
"""근거 설명문에서 특허명·논문명을 본문과 구분해 표시.

근거문은 35B 가 쓴 한 덩어리 텍스트라서 특허명·논문명이 일반 문장에 묻힌다. 어떤
특허를 이전받는지가 이 보고서의 핵심이므로, 실제 특허·논문 제목과 일치하는 구간만
찾아 별도 서식(진한 강조색)으로 렌더링한다.

찾는 방법
  · 그 과제의 **실제** 특허명·논문명 목록만 대상으로 한다(모델이 지어낸 이름은
    강조하지 않는다 — 강조가 곧 '검증된 인용'이라는 뜻이 되게 한다).
  · 공백·따옴표 차이를 무시하고 찾는다. 근거문은 '표면 영상유도 기반의 …' 처럼
    따옴표를 두르기도 하고 그냥 쓰기도 한다.
  · 긴 이름부터 찾아 짧은 이름이 긴 이름의 일부를 먹지 않게 한다.
  · 이름이 통째로 안 나오면 앞쪽 어절만 인용한 경우(예: 제목이 길어 앞부분만 씀)를
    위해 앞에서부터 어절을 줄여가며 다시 찾는다. 최소 길이(MIN_LEN) 아래로는 내려가지
    않는다 — '측정 장치' 같은 일반 어구까지 강조되면 오히려 읽기 어렵다.

표기 규칙(고정)
  · 강조 구간은 **반드시 꺽쇠(『 』)로 감싼다** — 원문에 인용부호가 있으면 흡수해서 바꾸고,
    없으면 새로 붙인다. 근거문은 같은 특허를 따옴표 있게/없게 섞어 쓰기 때문에(실측 59/135),
    표기를 여기서 통일하지 않으면 어떤 이름은 인용부호가 붙고 어떤 이름은 안 붙는다.
  · 강조는 **실제 성과 목록과 일치한 구간에만** 붙는다. 일반 설명 문장은 어떤 경우에도
    강조하지 않는다(verify_report_supply 가 산출물에서 다시 검사한다).

산출: [(문자열, 종류)] 목록. 종류는 "" (일반) / "특허" / "논문".
      종류가 있는 조각의 문자열은 이미 "『이름』" 형태로 꺽쇠가 붙어 있다.
"""
import re
import unicodedata

# 전체 제목이 그대로 나온 경우와, 앞어절만 인용된 경우의 기준을 분리한다.
# 전체 일치는 그 자체로 식별력이 있으므로 짧아도 인정한다('근감소증 진단 시스템' = 9자).
# 앞어절 일치는 일반 어구와 겹칠 위험이 커서 더 길고 문맥 증거까지 요구한다.
FULL_MIN_LEN = 6       # 전체 제목 일치 최소 길이(공백 제외)
MIN_LEN = 10           # 앞어절 부분일치 최소 길이
MIN_WORDS = 2          # 최소 어절 수 — 한 낱말 제목('책상')은 어떤 경우에도 강조하지 않는다
MIN_RATIO = 0.6        # 앞어절만 인용한 경우, 원 제목 길이의 이 비율 이상
# ---- 인용인가, 그냥 서술인가 -------------------------------------------------
# 제목과 문자열이 일치해도 그것이 **인용**이 아니라 서술의 일부일 수 있다. 실측 사례:
#   "관련 특허 1건이 존재하여 인체 호흡기관의 시뮬레이션 방법과 같은 기술적 기반이
#    확보되어 있다" — 제목이 맞지만 '…과 같은'으로 이어지는 서술이다. 이걸 강조하면
#   독자는 무엇이 인용인지 알 수 없다. 그래서 **모든 일치**에 인접 문맥 증거를 요구한다.
# 증거는 넷 중 하나(실측 103건의 근거 분포: 따옴표 59 · 뒤특허 31 · 앞라벨 8 · 연쇄 5).
CTX_AFTER = 16         # 뒤쪽 창(자)
CTX_BEFORE = 18        # 앞쪽 창(자)
_RE_AFTER = re.compile(r"특허|논문")
_RE_LABEL = re.compile(r"(특허|논문)[가-힣]{0,2}\s*['\"‘“]?\s*$")
_RE_CHAIN = re.compile(r"([,·]|및|와|과|또는)\s*['\"‘“]?\s*$")
# 산출 표기 — 꺽쇠(겹낫표). 본문 문장부호와 겹치지 않아 제목 경계가 또렷하다.
# 원문이 어떤 인용부호를 썼든( ' " ‘ ’ “ ” 「 」 〈 〉 등) 흡수한 뒤 이 표기로 통일한다.
QUOTE_L, QUOTE_R = "\u300E", "\u300F"       # 『 』
_Q = "'\"‘’“”「」『』《》〈〉"


def _fold(ch):
    """전각→반각 등 호환 정규화. 길이가 바뀌는 문자는 원문을 유지한다(색인 대응 보존)."""
    n = unicodedata.normalize("NFKC", ch)
    return n if len(n) == 1 else ch


def _norm(t):
    """비교용 정규화 — 공백·따옴표 제거 + 전각/반각 통일.

    nice_patent 의 특허명에는 전각 라틴 문자가 섞여 있다(예: '3Ｄ 간 오브젝트',
    '다상 간 ＣＴ 정합'). 모델은 이를 반각('3D', 'CT')으로 쓰기 때문에, 전각을
    접지 않으면 실제 인용을 놓친다 — 실측으로 확인된 누락 원인이다.
    """
    return "".join(_fold(c) for c in str(t).strip(_Q) if not c.isspace())


def _variants(name):
    """이름 → [(후보, 전체일치인가)]. 긴 것부터."""
    w = str(name).strip().strip(_Q).split()
    if len(w) < MIN_WORDS:
        return []
    full = len(_norm(" ".join(w)))
    out = []
    if full >= FULL_MIN_LEN:
        out.append((" ".join(w), True))
    for k in range(len(w) - 1, MIN_WORDS - 1, -1):
        cand = " ".join(w[:k])
        nl = len(_norm(cand))
        if nl >= MIN_LEN and nl >= full * MIN_RATIO:
            out.append((cand, False))
    return out


def _find(text, cand):
    """공백을 무시하고 text 안에서 cand 위치를 찾는다 → (start, end) 또는 None."""
    nc = _norm(cand)
    if not nc:
        return None
    # 원문 문자 인덱스 ↔ 정규화 문자열 인덱스 대응
    idx, buf = [], []
    for i, ch in enumerate(text):
        if ch.isspace() or ch in _Q:
            continue
        buf.append(_fold(ch))          # 비교 문자열도 같은 규칙으로 접는다
        idx.append(i)
    pos = "".join(buf).find(nc)
    if pos < 0:
        return None
    return idx[pos], idx[pos + len(nc) - 1] + 1


def cite_reason(text, a, b, prev_ok=False):
    """구간 [a,b) 가 인용인 근거 → 문자열, 아니면 "".

    ① 따옴표 : 원문이 이미 그 이름을 따옴표로 감쌌다 — 모델이 제목으로 표시한 것이다
    ② 뒤특허 : 뒤 CTX_AFTER 자 안에 '특허'/'논문'이 있다 ('… 특허는', '…과 같은 특허는')
    ③ 앞라벨 : 바로 앞이 라벨구다 ("관련 특허인 '", "관련 논문은 '")
    ④ 연쇄   : 앞이 나열 구분자이고 직전 구간이 인용으로 인정됐다 ("'A'와 'B'")
    """
    before = text[max(0, a - CTX_BEFORE):a]
    after = text[b:b + CTX_AFTER]
    lb = before.rstrip()[-1:]
    la = after.lstrip()[:1]
    if lb in _Q and la in _Q:
        return "따옴표"
    if _RE_AFTER.search(after):
        return "뒤특허"
    if _RE_LABEL.search(before):
        return "앞라벨"
    if prev_ok and _RE_CHAIN.search(before):
        return "연쇄"
    return ""


def _absorb_quotes(text, a, b):
    """이름 앞뒤에 이미 붙은 따옴표(와 그 사이 공백)를 구간에 포함시킨다.

    원문 표기를 그대로 두면 "'이름'" 과 "이름" 이 섞여서 보고서에 인용부호가 들쭉날쭉
    나온다. 흡수한 뒤 이 모듈이 정한 따옴표로 다시 감싸 표기를 통일한다.
    """
    i = a - 1
    while i >= 0 and text[i].isspace():
        i -= 1
    if i >= 0 and text[i] in _Q:
        a = i
    j = b
    while j < len(text) and text[j].isspace():
        j += 1
    if j < len(text) and text[j] in _Q:
        b = j + 1
    return a, b


def mark(text, patents=(), papers=()):
    """근거문 → [(조각, 종류)]. 종류: "" | "특허" | "논문"."""
    text = str(text or "")
    if not text:
        return []
    spans = []                                    # (start, end, kind)
    cands = ([(n, "특허") for n in patents if str(n).strip()]
             + [(n, "논문") for n in papers if str(n).strip()])
    cands.sort(key=lambda x: -len(_norm(x[0])))   # 긴 이름 먼저
    for name, kind in cands:
        for cand, is_full in _variants(name):
            hit = _find(text, cand)
            if not hit:
                continue
            a, b = hit
            if any(a < e and s < b for s, e, _ in spans):   # 이미 잡힌 구간과 겹침
                break
            spans.append((a, b, kind))
            break
    # 위치순으로 문맥을 판정한다(연쇄 근거는 앞 구간의 판정 결과를 쓴다).
    spans.sort()
    ok_spans, prev_ok = [], False
    for a, b, kind in spans:
        why = cite_reason(text, a, b, prev_ok)
        if why:
            ok_spans.append((a, b, kind))
            prev_ok = True
        else:
            prev_ok = False
    spans = ok_spans
    if not spans:
        return [(text, "")]
    spans.sort()
    out, cur = [], 0
    for a, b, kind in spans:
        a, b = _absorb_quotes(text, a, b)
        if a < cur:                      # 따옴표 흡수로 앞 구간과 겹치면 건너뛴다
            continue
        if a > cur:
            out.append((text[cur:a], ""))
        out.append((QUOTE_L + text[a:b].strip().strip(_Q) + QUOTE_R, kind))
        cur = b
    if cur < len(text):
        out.append((text[cur:], ""))
    return [(t, k) for t, k in out if t]


def names_of(pid, patents_map, papers_map):
    """과제고유번호 → (특허명 목록, 논문명 목록)."""
    pats = [x.get("특허명", "") for x in (patents_map or {}).get(str(pid), [])]
    paps = list((papers_map or {}).get(str(pid), []) or [])
    return pats, paps


# ---- 셀프테스트 -------------------------------------------------------------
# 회귀 방지용. verify_report_supply.py 가 보고서 검증 때마다 호출하고, 이 파일을 직접
# 실행해도 돌아간다. 여기 케이스는 실제로 났던 오류에서 뽑았다.
_CASES = [
    # (본문, 특허명, 논문명, 강조되어야 하는가, 메모)
    # ── 인용으로 인정하는 네 가지 근거
    ("이 과제가 확보한 '표면 영상유도 기반의 환자 위치 정렬 및 모니터링 시스템' 특허는 유용하다.",
     ["표면 영상유도 기반의 환자 위치 정렬 및 모니터링 시스템"], [], True,
     "①따옴표 — 원문이 제목으로 표시함"),
    ("과제는 방사선 치료 시뮬레이션 시스템 및 그 방법 특허를 확보했다.",
     ["방사선 치료 시뮬레이션 시스템 및 그 방법"], [], True,
     "②뒤특허 — 따옴표 없어도 뒤에 '특허'"),
    ("특히 '수술 검체용 분석 장치 및 분석 방법'과 같은 특허는 유용하다.",
     ["수술 검체용 분석 장치 및 분석 방법"], [], True,
     "②뒤특허 — '과 같은' 뒤에 '특허'가 오면 인용"),
    ("완료했다. 관련 특허인 근감소증 진단 시스템은 정밀하다.",
     ["근감소증 진단 시스템"], [], True, "③앞라벨 — '관련 특허인'"),
    ('논문 "TS-Net: A Deep Learning Framework" 를 보면',
     [], ["TS-Net: A Deep Learning Framework"], True, "①따옴표(쌍따옴표 흡수)"),
    # ── 인용이 아니어서 강조하면 안 되는 경우
    ("관련 특허 1건이 존재하여 인체 호흡기관의 시뮬레이션 방법과 같은 기술적 기반이 확보되어 있다.",
     ["인체 호흡기관의 시뮬레이션 방법"], [], False,
     "제목이 맞지만 '…과 같은 기술적 기반'으로 이어지는 서술 — 실제 오류 사례"),
    ("정밀 온도 제어 기술은 고주파 소작용 의료기기의 발열부 설계에 쓰인다.",
     ["고주파 소작용 의료기기"], [], False, "제목이 일반 명사처럼 쓰인 서술"),
    ("이 과제는 근감소증 진단 시스템을 개발한다.",
     ["근감소증 진단 시스템"], [], False, "인용 근거 없는 서술 — 강조 금지"),
    ("이 기업은 의료 영상을 분석하기 위한 소프트웨어를 만든다.",
     ["의료 영상을 분석하기 위한 방법"], [], False,
     "제목 앞머리와 겹치는 일반 설명 — 강조 금지"),
    ("호흡기 구조물 자동 분할 기술을 보유한다.",
     ["호흡기 구조물 자동 분할 방법 및 장치"], [], False, "앞어절 + 증거 없음"),
    ("사무용 책상을 만든다.", ["책상"], [], False, "한 낱말 제목 — 강조 금지"),
    ("측정 장치를 개발한다.", ["측정 장치"], [], False, "너무 짧은 일반 어구 — 강조 금지"),
]


def selftest():
    """→ 실패 목록(빈 목록이면 통과)."""
    bad = []
    for text, pats, paps, want, memo in _CASES:
        segs = mark(text, pats, paps)
        hits = [t for t, k in segs if k]
        if bool(hits) != want:
            bad.append(f"{memo}: 기대={'강조' if want else '없음'} 결과={hits}")
            continue
        for t in hits:                        # 표기 규칙: 항상 꺽쇠 한 겹
            if not (t.startswith(QUOTE_L) and t.endswith(QUOTE_R)):
                bad.append(f"{memo}: 따옴표 없음 {t!r}")
            elif t[1:-1].strip(_Q) != t[1:-1]:
                bad.append(f"{memo}: 따옴표 중복 {t!r}")
        if "".join(t for t, _ in segs).replace(QUOTE_L, "").replace(QUOTE_R, "") \
                .replace(" ", "") not in text.replace(" ", "").replace('"', "") \
                .replace("'", "").replace("‘", "").replace("’", ""):
            bad.append(f"{memo}: 본문 훼손")
    return bad


if __name__ == "__main__":
    fails = selftest()
    print(f"셀프테스트 {len(_CASES)}건 · 실패 {len(fails)}건")
    for f in fails:
        print("  -", f)
    raise SystemExit(1 if fails else 0)


# ---- 논문 표에 실을 목록 고르기 ---------------------------------------------
# 논문은 건수가 많아 표에 상한을 둔다. 그런데 근거문이 인용한 논문이 표에 없으면
# 독자가 강조된 제목을 같은 페이지에서 확인할 수 없다. 그래서
#   ① 최신 게재연도 상한(n_min)건  +  ② **어느 기업에서든** 인용된 논문 전건
# 을 합쳐 싣는다. ②를 기업별로 계산하면 같은 과제가 기업마다 다른 표를 보이게 되므로
# (앞서 특허에서 지적된 혼선) 반드시 전 기업 합집합으로 구한다.

def cited_paper_map(demands, patents_map, papers_map, top_key="top5"):
    """보고서 입력 JSON → {과제고유번호: {인용된 논문명(정규화), …}}.

    전 기업의 근거문을 훑어 그 과제에서 인용된 논문 제목을 모은다.
    """
    out = {}
    for v in (demands or {}).values():
        for t in (v.get(top_key) or v.get("top10") or []):
            pid = str(t.get("과제고유번호", ""))
            if not pid:
                continue
            pats = [x.get("특허명", "") for x in (patents_map or {}).get(pid, [])]
            paps = [x.get("논문명", "") for x in (papers_map or {}).get(pid, [])]
            if not paps:
                continue
            names = {_norm(x) for x in paps}
            for seg, kind in mark(t.get("추천근거_상세", ""), pats, paps):
                if kind != "논문":
                    continue
                n = _norm(seg)
                if n in names:
                    out.setdefault(pid, set()).add(n)
    return out


def pick_papers(paps, cited=(), n_min=10):
    """표에 실을 논문 목록 → (목록, 전체 건수).

    최신 n_min 건 + 인용된 논문 전건. 원래 순서(최신순)를 유지한다.
    """
    paps = list(paps or [])
    if not paps:
        return [], 0
    cited = set(cited or ())
    keep = set(range(min(n_min, len(paps))))
    for i, x in enumerate(paps):
        if _norm(x.get("논문명", "")) in cited:
            keep.add(i)
    return [paps[i] for i in sorted(keep)], len(paps)
