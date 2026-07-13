# -*- coding: utf-8 -*-
"""기존 산출물의 '마침표 뒤 띄어쓰기 누락' 직접 교정(LLM 미사용).
필터 pkl 의 추천근거_상세·LLM판단근거에 compa_match.normalize_spacing 적용 →
이후 _build_full_inputs.py 로 통합best.json 재생성."""
import pandas as pd
from compa_match import normalize_spacing

PKLS = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]
COLS = ["추천근거_상세", "LLM판단근거"]

for f in PKLS:
    df = pd.read_pickle(f)
    changed = 0
    for c in COLS:
        if c not in df.columns:
            continue
        new = df[c].apply(lambda s: normalize_spacing(s) if isinstance(s, str) else s)
        changed += int((new != df[c]).sum())
        df[c] = new
    df.to_pickle(f)
    df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
    print(f"{f}: {changed}개 셀 교정")
print("완료")
