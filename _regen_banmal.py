# -*- coding: utf-8 -*-
"""상세 매칭 근거(추천근거_상세)를 Qwen 으로 반말(한다체)로 재생성해 교체.
의미·수치·고유명사·[섹션 제목]·줄바꿈은 그대로 두고 종결어미만 한다체로 변환.
항목별 체크포인트(중단·재개). 완료 후 _build_full_inputs.py 로 통합best.json 재생성 필요."""
import json, os, re, time
import pandas as pd
import compa_match as cm

S = os.environ.get("COMPA_SCRATCH", "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
CK = f"{S}/banmal_ckpt.json"
PKLS = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]

SYS = ("너는 한국어 문체 교정기다. 입력 텍스트의 의미·정보·수치·고유명사·[대괄호 섹션 제목]·줄바꿈 구조를 "
       "100% 그대로 유지하고, 오직 문장 종결 문체만 '~다/~한다/~이다/~된다/~있다' 형태의 평서형(한다체·반말)으로 "
       "바꾼다. '~습니다/~입니다/~됩니다/~있습니다/~합니다/~됩니다/~집니다' 등 존댓말 종결은 모두 반말로 바꾼다. "
       "내용 추가·삭제·요약·재배열 금지. 설명이나 머리말 없이 변환된 본문만 출력한다.")
JON = re.compile(r"(습니다|입니다|됩니다|있습니다|합니다|봅니다|립니다|칩니다|십니다|믿습니다|납니다)")

def log(*a): print(*a, flush=True)

def convert(text):
    msgs = [{"role": "system", "content": SYS}, {"role": "user", "content": text}]
    out = cm.stream_explanation(msgs, max_tokens=1800, temperature=0.0, top_p=1.0)
    out = out.strip()
    out = re.sub(r"^(변환\s*결과|출력)\s*[:：]\s*", "", out)   # 혹시 붙는 머리말 제거
    return cm.normalize_spacing(out)

def main():
    ck = json.load(open(CK)) if os.path.exists(CK) else {}
    frames = {f: pd.read_pickle(f) for f in PKLS}
    todo = [(f, i, r) for f, df in frames.items() for i, r in df.iterrows()]
    log(f"대상 {len(todo)}건 (체크포인트 {len(ck)})")
    cm.load_model_blocking(progress_cb=lambda m: log("  " + m))
    n = 0
    for f, i, r in todo:
        n += 1
        key = f"{r['번호']}::{r['과제고유번호']}"
        src = str(r["추천근거_상세"] or "")
        if key in ck:
            continue
        if not JON.search(src):                 # 이미 존댓말 없음 → 그대로
            ck[key] = src
        else:
            ts = time.time()
            new = convert(src)
            # 안전장치: 길이 급감/빈 출력/섹션 소실 시 원문 유지
            secs_ok = src.count("[") == new.count("[")
            if new and len(new) > len(src) * 0.6 and secs_ok and not JON.search(new):
                ck[key] = new
                log(f"[{n}/{len(todo)}] {key} 변환 ({time.time()-ts:.0f}s)")
            else:
                ck[key] = src
                log(f"[{n}/{len(todo)}] {key} 유지(검증실패: len={len(new)} secs_ok={secs_ok} 존댓말잔존={bool(JON.search(new))})")
        if n % 5 == 0:
            json.dump(ck, open(CK, "w"), ensure_ascii=False)
    json.dump(ck, open(CK, "w"), ensure_ascii=False)

    # pkl 반영
    for f, df in frames.items():
        df["추천근거_상세"] = [ck.get(f"{r['번호']}::{r['과제고유번호']}", r["추천근거_상세"])
                          for _, r in df.iterrows()]
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
        log(f"저장 {f}")
    left = sum(1 for _, r in pd.concat(frames.values()).iterrows() if JON.search(str(r["추천근거_상세"] or "")))
    log(f"완료. 존댓말 잔존 행: {left}")

if __name__ == "__main__":
    main()
