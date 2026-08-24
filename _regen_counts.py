# -*- coding: utf-8 -*-
"""근거문의 실적 건수가 실적 목록과 어긋난 문장만 35B 로 다시 써서 교체.

프롬프트는 patent_list_count·paper_list_count 로 건수를 주는데(compa_match:719-720),
모델이 그 값 대신 '인용 특허 목록의 길이'를 세어 적는 사례가 있다. 실측 v33 에서
3개 수치(2문장)가 특허·논문 실적 표와 어긋났다:

  · 엠비디(주) TOP5(1711137934)  "논문 1건과 특허 1건" → 실제 특허 2건 · 논문 0건
  · 세로메드(주) TOP1(1465032903) "논문 2건과 특허 1건" → 실제 특허 2건

문장 전체를 새로 쓰지 않고 해당 '문장 하나'만 모델에게 교정시킨다(다른 서술·인용
제목·문체가 함께 바뀌면 검토가 다시 필요해진다). 텍스트는 전부 모델이 쓴다.

반영: explain 체크포인트 · 매칭 pkl/xlsx · 보고서 JSON (네 곳을 같이 고쳐야 다음
재생성에서 되살아나지 않는다). 이후 docx·pdf 재생성 필요.

사용: COMPA_SCRATCH=<scratch> python _regen_counts.py [--dry-run]
"""
import argparse
import json
import os
import re

import pandas as pd

import compa_match as cm
import explain_meditek_top10 as xp
import match_meditek_supply as ms

SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
TAG = "MEDITEK_260820"
JSON_PATH = f"{TAG}_보고서.json"
PKL_PATH = f"{TAG}_매칭.pkl"
XLSX_PATH = f"{TAG}_매칭.xlsx"
EX_CKPT = f"supply_{TAG}_explain_v5_ckpt.json"
FIELDS = ("판단근거", "추천근거_상세")
RETRY = 3

_CNT = re.compile(r"(특허|논문)\s*(\d+)\s*건")
# 기업이 보유한 특허 건수를 인용하는 문장은 과제 실적과 무관하다(문맥으로 제외).
_OWN = re.compile(r"당사|자사|이 기업|보유")
_JON = re.compile(r"습니다|입니다|됩니다|합니다")

SYS = ("너는 한국어 문장 교정기다. 입력 문장에 적힌 '과제의 연구성과 건수'를 주어진 실제 "
       "건수에 맞게 고쳐 문장 하나로 다시 쓴다.\n"
       "규칙:\n"
       "1) 특허·논문 건수는 주어진 실제 건수만 숫자로 인용한다.\n"
       "2) 실제 건수가 0인 실적은 문장에서 아예 언급하지 않는다(그 실적을 뺀 나머지 내용만 "
       "남긴다). 0건이라고 적지도 않는다.\n"
       "3) 건수 외의 정보·고유명사·인용된 제목·어순·표현은 최대한 그대로 두고, 문체는 "
       "'~다/~한다' 평서형(한다체)을 유지한다.\n"
       "4) 새로운 사실·평가·수치를 만들지 않는다.\n"
       "5) 설명·머리말 없이 교정된 문장 하나만 출력한다.")


def log(*a):
    print(*a, flush=True)


def counts_of(pid, pat, pap):
    return {"특허": len(pat.get(str(pid), [])), "논문": len(pap.get(str(pid), []))}


def mismatches(text, real):
    """(match, 실제건수) 목록 — 실적 목록과 어긋난 건수 표기."""
    out = []
    for m in _CNT.finditer(text or ""):
        if _OWN.search(text[max(0, m.start() - 12):m.start()]):
            continue
        if int(m.group(2)) != real[m.group(1)]:
            out.append((m, real[m.group(1)]))
    return out


def replace_loose(text, old, new):
    """old 를 new 로 교체 — 공백 차이는 무시한다.

    체크포인트에는 모델 원문이 그대로 들어 있고('논문 2 건'), 보고서 JSON·pkl 에는
    report_text_fix 로 교정된 문장이 들어 있다('논문 2건'). 같은 문장을 양쪽에서
    갈아끼우려면 공백을 느슨하게 봐야 한다.
    """
    if old in text:
        return text.replace(old, new), True
    pat = r"[ \t]*".join(re.escape(c) for c in re.sub(r"\s+", "", old))
    m = re.search(pat, text)
    return (text[:m.start()] + new + text[m.end():], True) if m else (text, False)


def sentence_span(text, i):
    """i 를 포함하는 문장 구간 [a, b). 섹션 라벨('[연관성] ')·줄바꿈도 경계로 본다."""
    starts = [0] + [m.end() for m in re.finditer(r"\.\s+|\n+|\]\s+", text)]
    ends = [m.start() + 1 for m in re.finditer(r"\.(?=\s|$)", text)] + [len(text)]
    return max(s for s in starts if s <= i), min(e for e in ends if e > i)


def fix_sentence(sent, real, retry=RETRY):
    """문장 하나 → 건수를 바로잡은 문장. 검증 실패 시 빈 문자열."""
    payload = {"문장": sent, "실제_특허건수": real["특허"], "실제_논문건수": real["논문"]}
    msgs = [{"role": "system", "content": SYS},
            {"role": "user", "content": json.dumps(payload, ensure_ascii=False)}]
    for i in range(retry):
        out = cm.stream_explanation(msgs, max_tokens=400,
                                    temperature=0.0 if i == 0 else 0.2, top_p=1.0)
        new = cm.normalize_spacing(re.sub(r"^(교정\s*결과|출력|문장)\s*[:：]\s*", "",
                                          out.strip()).strip().strip('"'))
        why = validate(new, sent, real)
        if not why:
            return new
        log(f"      재시도 {i + 1}/{retry} — {why}")
    return ""


def validate(new, old, real):
    """교정문 검증 → 실패 이유(통과면 "")."""
    if not new:
        return "빈 출력"
    if not (len(old) * 0.4 <= len(new) <= len(old) * 1.8):
        return f"길이 이상({len(new)} vs 원문 {len(old)})"
    if "[" in new or "\n" in new:
        return "섹션 라벨·줄바꿈 유입"
    if _JON.search(new):
        return "존댓말 종결"
    if xp.BANNED.search(new):
        return "금지 표현(매칭 기준 노출)"
    for m in _CNT.finditer(new):
        if _OWN.search(new[max(0, m.start() - 12):m.start()]):
            continue
        if int(m.group(2)) != real[m.group(1)]:
            return f"건수 여전히 불일치({m.group(0)} ≠ {real[m.group(1)]}건)"
    for kind in ("특허", "논문"):
        if real[kind] == 0 and re.search(kind + r"\s*\d+\s*건", new):
            return f"{kind} 0건인데 건수 인용"
    return ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="대상만 찾고 생성하지 않는다")
    a = ap.parse_args()

    J = json.load(open(JSON_PATH, encoding="utf-8"))
    PAT = json.load(open(os.path.join(SCRATCH, "pid_patents.json"), encoding="utf-8"))
    _pp = os.path.join(SCRATCH, "pid_papers.json")
    PAP = json.load(open(_pp, encoding="utf-8")) if os.path.exists(_pp) else {}

    todo = []                       # (기업명, 번호, rank, pid, 필드, 문장, 실제건수)
    for key, v in J.items():
        for t in v["top5"]:
            pid = str(t["과제고유번호"])
            real = counts_of(pid, PAT, PAP)
            for f in FIELDS:
                text = t.get(f, "") or ""
                spans = {sentence_span(text, m.start()) for m, _ in mismatches(text, real)}
                for a_, b_ in sorted(spans):
                    todo.append((v["기업명"], key, t["rank"], pid, f,
                                 text[a_:b_].strip(), real))
    log(f"건수 불일치 문장 {len(todo)}건")
    for x in todo:
        log(f"  · {x[0]} TOP{x[2]} [{x[4]}] 실제 특허 {x[6]['특허']}건 · 논문 {x[6]['논문']}건")
        log(f"      원문: {x[5]}")
    if not todo or a.dry_run:
        return

    cm.load_model_blocking(progress_cb=lambda m: log("  " + m))
    fixed = []                      # (기업명, key, rank, pid, 필드, 원문장, 새문장)
    for nm, key, rank, pid, f, sent, real in todo:
        log(f"\n[{nm} TOP{rank} {f}] 교정 중…")
        new = fix_sentence(sent, real)
        if not new:
            log("      실패 — 원문 유지")
            continue
        log(f"      교정: {new}")
        fixed.append((nm, key, rank, pid, f, sent, new))
    if not fixed:
        log("교체할 문장 없음")
        return

    # ---- 반영 ①  보고서 JSON
    for nm, key, rank, pid, f, old, new in fixed:
        for t in J[key]["top5"]:
            if t["rank"] == rank:
                t[f] = xp.polish(t[f].replace(old, new))
    json.dump(J, open(JSON_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    log(f"\n저장 {JSON_PATH}")

    # ---- 반영 ②  매칭 pkl · xlsx
    df = pd.read_pickle(PKL_PATH)
    for nm, key, rank, pid, f, old, new in fixed:
        sel = (df["기업명"] == nm) & (df["rank"] == rank) & (df["과제고유번호"].astype(str) == pid)
        if not sel.any():
            log(f"  (주의) pkl 행 없음: {nm} TOP{rank}")
            continue
        df.loc[sel, f] = df.loc[sel, f].astype(str).str.replace(old, new, regex=False)
    df.to_pickle(PKL_PATH)
    ms.save_xlsx(df, XLSX_PATH)
    log(f"저장 {PKL_PATH} · {XLSX_PATH}")

    # ---- 반영 ③  explain 체크포인트(다음 재생성에서 되살아나지 않게)
    if os.path.exists(EX_CKPT):
        ck = json.load(open(EX_CKPT, encoding="utf-8"))
        n = 0
        for nm, key, rank, pid, f, old, new in fixed:
            if f != "추천근거_상세":
                continue
            no = str(df.loc[(df["기업명"] == nm) & (df["rank"] == rank), "번호"].iloc[0])
            ckey = f"{no}::{pid}"
            if ckey not in ck:
                log(f"  (주의) 체크포인트에 항목 없음: {ckey}")
                continue
            ck[ckey], ok = replace_loose(ck[ckey], old, new)
            n += ok
            if not ok:
                log(f"  (주의) 체크포인트에 원문장 없음: {ckey}")
        json.dump(ck, open(EX_CKPT, "w", encoding="utf-8"), ensure_ascii=False)
        log(f"저장 {EX_CKPT} ({n}건 반영)")

    # ---- 잔존 확인
    left = 0
    for v in J.values():
        for t in v["top5"]:
            real = counts_of(t["과제고유번호"], PAT, PAP)
            left += sum(len(mismatches(t.get(f, "") or "", real)) for f in FIELDS)
    log(f"\n완료. 남은 건수 불일치 {left}건 — docx·pdf 재생성 필요")


if __name__ == "__main__":
    main()
