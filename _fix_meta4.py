# -*- coding: utf-8 -*-
"""잔존 메타/회피 [유사 사례 및 실적] 소수 건: 실적 유무 언급 자체를 금지하고 기술 연결성만 서술."""
import re, pickle, json
import pandas as pd
import compa_match as cm

TARGET = [("COMPA_필터11_20_최종추천.pkl", "16", "1345327373"),
          ("COMPA_필터11_20_최종추천.pkl", "20", "1345367592"),
          ("COMPA_필터21_78_최종추천.pkl", "24", "1711106063"),
          ("COMPA_필터21_78_최종추천.pkl", "48", "2540000128")]
META = re.compile(r"이 섹션은|설명을 마무리|생략(한다|합니다)|서술할 수 없|설명할 수 없|분석할 수 없|제공된 정보에는|제공되지 않|실적이 없|포함하지 않는다|언급.{0,4}생략")
def norm(s): return re.sub(r"\s+", "", str(s or ""))
jb = json.load(open("COMPA_통합best.json", encoding="utf-8"))
dm = {(norm(e["기업명"]), norm(e["수요기술명"])): (e.get("수요기술 내용", ""), e.get("수요기술 사양", "")) for e in jb.values()}

SYS = ("너는 기술이전 보고서의 '유사 사례 및 실적' 문단을 쓴다. 이 과제는 인용할 논문/특허가 없으므로 "
       "논문·특허·실적의 유무를 절대 언급하지 마('없다/제공되지 않았다/생략한다/설명할 수 없다' 등 일절 금지). "
       "대신 과제가 개발·규명하는 기술 내용·방법이 기업 수요 해결에 실제로 어떻게 활용·연결되는지만 2~3문장으로 "
       "구체적으로 서술한다. 반드시 반말 한다체(~다/~한다/~이다), 존댓말·메타 문구 금지. 머리말 없이 문단만 출력한다.")

def gen(과제명, 설명, 수요명, 내용):
    u = (f"[과제] 과제명: {과제명}\n설명: {설명[:600]}\n\n[기업 수요] 수요기술명: {수요명}\n내용: {내용[:300]}\n\n"
         "'유사 사례 및 실적' 문단(실적 유무 언급 금지, 기술 연결성만):")
    for temp in (0.3, 0.6, 0.9):
        o = cm.normalize_spacing(cm.stream_explanation(
            [{"role": "system", "content": SYS}, {"role": "user", "content": u}],
            max_tokens=380, temperature=temp, top_p=0.9).strip())
        if o and not META.search(o) and not re.search(r"습니다|입니다|됩니다", o):
            return o
    return None

def main():
    frames = {}
    cm.load_model_blocking(progress_cb=lambda m: print("  " + m, flush=True))
    for f, no, pid in TARGET:
        df = frames.setdefault(f, pd.read_pickle(f))
        idx = next(i for i, r in df.iterrows() if str(r["번호"]) == no and str(r["과제고유번호"]) == pid)
        r = df.loc[idx]
        내용, _ = dm.get((norm(r["기업명"]), norm(r["수요기술명"])), ("", ""))
        new = gen(r["과제명"], str(r["과제설명문"] or ""), r["수요기술명"], 내용)
        if not new:
            print(f"  {no}::{pid} 실패(유지)", flush=True); continue
        parts = re.split(r"(\[[^\]]+\])", str(r["추천근거_상세"]))
        for j in range(1, len(parts), 2):
            if "유사 사례" in parts[j]:
                parts[j + 1] = " " + new
        df.at[idx, "추천근거_상세"] = cm.normalize_spacing("".join(parts))
        print(f"  {no}::{pid} OK: {new[:60]}", flush=True)
    for f, df in frames.items():
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
    left = sum(1 for f, df in frames.items() for _, r in df.iterrows() if META.search(str(r["추천근거_상세"] or "")))
    print("완료. 잔존:", left, flush=True)

if __name__ == "__main__":
    main()
