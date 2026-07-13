# -*- coding: utf-8 -*-
"""내용 없는 회피형 [유사 사례 및 실적] 문장(punt_targets.json)을 기술 연결성 중심 반말로 재생성."""
import re, json
import pandas as pd
import compa_match as cm

S = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"
TARGET = json.load(open(f"{S}/punt_targets.json", encoding="utf-8"))
BAD = re.compile(r"제공되지 않|실적이 없|성과가 없|확인되지 않|작성하지 않|작성할 수 없|서술하지 않|서술할 수 없|"
                 r"제시할 수 없|포함할 수 없|언급할 수 없|생략|이 섹션|습니다|입니다|됩니다")
def norm(s): return re.sub(r"\s+", "", str(s or ""))
jb = json.load(open("COMPA_통합best.json", encoding="utf-8"))
dm = {(norm(e["기업명"]), norm(e["수요기술명"])): e.get("수요기술 내용", "") for e in jb.values()}

SYS = ("너는 기술이전 보고서의 '유사 사례 및 실적' 문단을 쓴다. 논문·특허·실적의 유무를 절대 언급하지 마"
       "('없다/제공되지 않았다/생략/작성하지 않는다/확인되지 않는다' 등 일절 금지). "
       "과제가 개발·규명하는 기술 내용·방법이 기업 수요 해결에 어떻게 활용·연결되는지만 2~3문장으로 구체적으로 "
       "서술한다. 반드시 반말 한다체, 존댓말·메타 문구 금지. 머리말 없이 문단만 출력한다.")

def gen(과제명, 설명, 수요명, 내용):
    u = (f"[과제] 과제명: {과제명}\n설명: {설명[:600]}\n\n[기업 수요] 수요기술명: {수요명}\n내용: {내용[:300]}\n\n"
         "'유사 사례 및 실적' 문단(실적 유무 언급 금지, 기술 연결성만):")
    for temp in (0.3, 0.6, 0.9, 0.9):
        o = cm.normalize_spacing(cm.stream_explanation(
            [{"role": "system", "content": SYS}, {"role": "user", "content": u}],
            max_tokens=380, temperature=temp, top_p=0.9).strip())
        if o and not BAD.search(o):
            return o
    return None

def main():
    frames = {}
    cm.load_model_blocking(progress_cb=lambda m: print("  " + m, flush=True))
    okn = 0
    for f, no, pid, sec in TARGET:
        df = frames.setdefault(f, pd.read_pickle(f))
        idx = next(i for i, r in df.iterrows() if str(r["번호"]) == no and str(r["과제고유번호"]) == pid)
        r = df.loc[idx]; 내용 = dm.get((norm(r["기업명"]), norm(r["수요기술명"])), "")
        new = gen(r["과제명"], str(r["과제설명문"] or ""), r["수요기술명"], 내용)
        if not new:
            print(f"  {no}::{pid} 실패(유지)", flush=True); continue
        parts = re.split(r"(\[[^\]]+\])", str(r["추천근거_상세"]))
        for j in range(1, len(parts), 2):
            if "유사 사례" in parts[j]:
                parts[j + 1] = " " + new
        df.at[idx, "추천근거_상세"] = cm.normalize_spacing("".join(parts)); okn += 1
        print(f"  {no}::{pid} OK: {new[:60]}", flush=True)
    for f, df in frames.items():
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
    print(f"완료. 성공 {okn}/{len(TARGET)}", flush=True)

if __name__ == "__main__":
    main()
