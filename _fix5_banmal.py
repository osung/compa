# -*- coding: utf-8 -*-
"""반말 변환에서 존댓말이 잔존해 원문 유지된 소수 행만 재시도(존댓말 0 결과 채택)."""
import re, os, json
import pandas as pd
import compa_match as cm

PKLS = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]
JON = re.compile(r"습니다|입니다|됩니다|있습니다|합니다|봅니다|집니다|립니다|칩니다|십니다")
SYS = ("너는 한국어 문체 교정기다. 입력 텍스트의 의미·수치·고유명사·[대괄호 섹션 제목]·줄바꿈 구조를 그대로 두고, "
       "모든 문장 종결을 예외 없이 '~다/~한다/~이다/~된다/~있다' 평서형(한다체)으로 바꾼다. "
       "'~습니다/~입니다/~됩니다/~있습니다/~합니다' 등 존댓말이 단 하나도 남지 않게 한다. "
       "내용 추가·삭제·요약 금지. 머리말 없이 본문만 출력한다.")

def convert(text, temp):
    out = cm.stream_explanation([{"role": "system", "content": SYS}, {"role": "user", "content": text}],
                                max_tokens=1900, temperature=temp, top_p=1.0).strip()
    out = re.sub(r"^(변환\s*결과|출력)\s*[:：]\s*", "", out)
    return cm.normalize_spacing(out)

def main():
    frames = {f: pd.read_pickle(f) for f in PKLS}
    targets = [(f, i) for f, df in frames.items() for i, r in df.iterrows()
               if JON.search(str(r["추천근거_상세"] or ""))]
    print("재시도 대상:", len(targets), flush=True)
    cm.load_model_blocking(progress_cb=lambda m: print(" ", m, flush=True))
    for f, i in targets:
        df = frames[f]; src = str(df.at[i, "추천근거_상세"] or "")
        best = None
        for temp in (0.0, 0.3, 0.6):
            new = convert(src, temp)
            if new and abs(new.count("[") - src.count("[")) == 0 and len(new) > len(src) * 0.6:
                if not JON.search(new):
                    best = new; break
                if best is None or len(JON.findall(new)) < len(JON.findall(best)):
                    best = new
        if best and len(best) > len(src) * 0.6:
            df.at[i, "추천근거_상세"] = best
            print(f"  {df.at[i,'번호']}::{df.at[i,'과제고유번호']} → 존댓말 {len(JON.findall(best))}개", flush=True)
    for f, df in frames.items():
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
    left = sum(1 for f, df in frames.items() for _, r in df.iterrows() if JON.search(str(r["추천근거_상세"] or "")))
    print("완료. 존댓말 잔존 행:", left, flush=True)

if __name__ == "__main__":
    main()
