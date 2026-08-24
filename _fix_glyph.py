# -*- coding: utf-8 -*-
"""근거문에 섞인 U+FFFD(치환 문자)를 복원해 교체.

35B 스트리밍 출력에서 한 글자가 U+FFFD 로 깨져 그대로 보고서까지 실렸다
(v33·v34 (주)위드세이브 TOP1 '바이�프린팅' 2곳 → PDF 7쪽에서 notdef 상자로 렌더).
2차 피해로 특허명 표기도 어긋난다 — 이름이 실적 목록과 한 글자 달라 meditek_cite_mark
가 인용으로 잡지 못하고, 『 』 대신 모델이 쓴 작은따옴표가 그대로 남는다.

깨진 글자는 지어내지 않고 **유일하게 결정될 때만** 복원한다:
  ① 그 과제의 특허명·논문명 목록(= 보고서 실적 표와 같은 소스)
  ② 같은 근거문의 다른 자리에 온전히 남아 있는 같은 낱말
후보 글자가 둘 이상이거나 하나도 없으면 손대지 않고 사람이 볼 목록으로 남긴다.

반영: 보고서 JSON · 매칭 pkl/xlsx · explain 체크포인트. 이후 docx·pdf 재생성 필요.

사용: COMPA_SCRATCH=<scratch> python _fix_glyph.py [--dry-run]
"""
import argparse
import json
import os
import re

import pandas as pd

import _regen_counts as rc
import match_meditek_supply as ms

SCRATCH = os.environ.get("COMPA_SCRATCH", ".")
TAG = "MEDITEK_260820"
JSON_PATH = f"{TAG}_보고서.json"
PKL_PATH = f"{TAG}_매칭.pkl"
XLSX_PATH = f"{TAG}_매칭.xlsx"
EX_CKPT = f"supply_{TAG}_explain_v5_ckpt.json"
RS_CKPT = f"supply_{TAG}_reason_v5_ckpt.json"
FIELDS = ("판단근거", "추천근거_상세")
BAD = "�"
# 낱말 경계 — 여기까지가 한 낱말이다(따옴표·괄호·구두점·공백)
_EDGE = re.compile(r"[\s'\"‘’“”『』「」()\[\]{}<>,.·:;!?/|]")


def log(*a):
    print(*a, flush=True)


def word_at(text, i):
    """i 를 포함하는 낱말 구간 [a, b)."""
    a = i
    while a > 0 and not _EDGE.match(text[a - 1]):
        a -= 1
    b = i + 1
    while b < len(text) and not _EDGE.match(text[b]):
        b += 1
    return a, b


def recover(word, sources):
    """깨진 낱말 → 복원 글자 후보 집합. word 는 BAD 를 정확히 하나 포함한다."""
    pat = re.compile("".join("." if c == BAD else re.escape(c) for c in word))
    found = set()
    for src in sources:
        for m in pat.finditer(src or ""):
            ch = m.group(0)[word.index(BAD)]
            if ch != BAD and not _EDGE.match(ch):
                found.add(ch)
    return found


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    J = json.load(open(JSON_PATH, encoding="utf-8"))
    PAT = json.load(open(os.path.join(SCRATCH, "pid_patents.json"), encoding="utf-8"))
    _pp = os.path.join(SCRATCH, "pid_papers.json")
    PAP = json.load(open(_pp, encoding="utf-8")) if os.path.exists(_pp) else {}

    todo, hopeless = [], []          # (기업명, key, rank, pid, 필드, 깨진낱말, 복원낱말)
    for key, v in J.items():
        for t in v["top5"]:
            pid = str(t["과제고유번호"])
            names = ([x.get("특허명", "") for x in PAT.get(pid, [])]
                     + [x.get("논문명", "") for x in PAP.get(pid, [])])
            for f in FIELDS:
                text = t.get(f, "") or ""
                seen = set()
                for m in re.finditer(BAD, text):
                    wa, wb = word_at(text, m.start())
                    word = text[wa:wb]
                    if word in seen:
                        continue
                    seen.add(word)
                    # 같은 근거문의 온전한 낱말도 근거로 쓴다(BAD 가 있는 자리는 제외됨)
                    cand = recover(word, names + [text])
                    if len(cand) == 1:
                        todo.append((v["기업명"], key, t["rank"], pid, f, word,
                                     word.replace(BAD, cand.pop())))
                    else:
                        hopeless.append((v["기업명"], t["rank"], f, word, sorted(cand)))

    log(f"복원 대상 낱말 {len(todo)}건 · 판단 보류 {len(hopeless)}건")
    for x in todo:
        log(f"  · {x[0]} TOP{x[2]} [{x[4]}]  {x[5]} → {x[6]}")
    for x in hopeless:
        log(f"  (보류) {x[0]} TOP{x[1]} [{x[2]}]  {x[3]}  후보 {x[4]}")
    if not todo or a.dry_run:
        return

    # ---- 반영 ①  보고서 JSON
    for nm, key, rank, pid, f, old, new in todo:
        for t in J[key]["top5"]:
            if t["rank"] == rank:
                t[f] = t[f].replace(old, new)
    json.dump(J, open(JSON_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    log(f"\n저장 {JSON_PATH}")

    # ---- 반영 ②  매칭 pkl · xlsx
    df = pd.read_pickle(PKL_PATH)
    for nm, key, rank, pid, f, old, new in todo:
        sel = (df["기업명"] == nm) & (df["rank"] == rank) & (df["과제고유번호"].astype(str) == pid)
        if not sel.any():
            log(f"  (주의) pkl 행 없음: {nm} TOP{rank}")
            continue
        df.loc[sel, f] = df.loc[sel, f].astype(str).str.replace(old, new, regex=False)
    df.to_pickle(PKL_PATH)
    ms.save_xlsx(df, XLSX_PATH)
    log(f"저장 {PKL_PATH} · {XLSX_PATH}")

    # ---- 반영 ③  체크포인트(다음 재생성에서 되살아나지 않게)
    for path, field in ((EX_CKPT, "추천근거_상세"), (RS_CKPT, "판단근거")):
        if not os.path.exists(path):
            continue
        ck = json.load(open(path, encoding="utf-8"))
        n = 0
        for nm, key, rank, pid, f, old, new in todo:
            if f != field:
                continue
            no = str(df.loc[(df["기업명"] == nm) & (df["rank"] == rank), "번호"].iloc[0])
            ckey = f"{no}::{pid}"
            if ckey not in ck:
                log(f"  (주의) {path} 에 항목 없음: {ckey}")
                continue
            ck[ckey], ok = rc.replace_loose(ck[ckey], old, new)
            n += ok
            if not ok:
                log(f"  (주의) {path} 에 원문 낱말 없음: {ckey} {old}")
        if n:
            json.dump(ck, open(path, "w", encoding="utf-8"), ensure_ascii=False)
            log(f"저장 {path} ({n}건 반영)")

    left = sum((t.get(f, "") or "").count(BAD) for v in J.values()
               for t in v["top5"] for f in FIELDS)
    log(f"\n완료. 남은 U+FFFD {left}자 — docx·pdf 재생성 필요")


if __name__ == "__main__":
    main()
