# -*- coding: utf-8 -*-
"""필터 재매칭 pkl 의 부정 톤 '매칭 근거(LLM판단근거)'만 골라 긍정형으로 재작성.
compa_match.gen_match_reason(긍정 few-shot) 사용. 이후 _build_full_inputs.py 로 통합best.json 재생성."""
import re, glob, json
import pandas as pd
import compa_match as cm

NEG = re.compile(r"부재|미포함|불일치|미흡|미달|부족|다름|어렵|불가능")
PKLS = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]

# (기업,수요) → (내용, 사양)  — pkl엔 없어서 통합best.json 에서 보강
def norm(s): return re.sub(r"\s+", "", str(s or ""))
jb = json.load(open("COMPA_통합best.json", encoding="utf-8"))
dem_meta = {(norm(e["기업명"]), norm(e["수요기술명"])): (e.get("수요기술 내용", ""), e.get("수요기술 사양", ""))
            for e in jb.values()}

cm.load_model_blocking(progress_cb=lambda m: print(" ", m, flush=True))

total = fixed = 0
for f in PKLS:
    df = pd.read_pickle(f)
    changed = 0
    for idx, r in df.iterrows():
        jg = str(r["LLM판단근거"] or "")
        if not NEG.search(jg):
            continue
        total += 1
        내용, 사양 = dem_meta.get((norm(r["기업명"]), norm(r["수요기술명"])), ("", ""))
        demand = {"수요기술명": r["수요기술명"], "수요기술 내용": 내용, "수요기술 사양": 사양}
        proj = {"과제명": r["과제명"], "설명": r.get("과제설명문", "")}
        new = cm.gen_match_reason(demand, proj)
        if new and not NEG.search(new):
            df.at[idx, "LLM판단근거"] = new; changed += 1; fixed += 1
            print(f"  [{f}] {r['과제명'][:22]}\n     - old: {jg}\n     + new: {new}", flush=True)
        else:
            print(f"  [{f}] ! 재작성 실패/여전히 부정 → 유지: {jg[:30]} (gen='{new[:30]}')", flush=True)
    if changed:
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
        print(f"  → {f} 저장({changed}건 수정)", flush=True)

print(f"\n대상 {total} / 수정 {fixed}")
