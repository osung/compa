# -*- coding: utf-8 -*-
"""특정 과제의 부자연스러운 [유사 사례 및 실적] 섹션을 자연스러운 반말 문장으로 재생성·교체."""
import re, json, pickle
import pandas as pd
import compa_match as cm

TARGET = [("COMPA_필터10_최종추천.pkl", "7", "1545023701")]  # (pkl, 번호, 과제고유번호)
SYS = ("너는 기술이전 보고서의 '유사 사례 및 실적' 문단을 쓰는 한국어 AI다. 주어진 과제와 기업 수요를 바탕으로, "
       "과제가 축적한 연구 노하우·기술적 성과가 기업 수요 해결에 주는 시사점을 2~3문장으로 자연스럽게 서술한다. "
       "규칙: (1) 반드시 반말 한다체(~다/~한다/~이다/~된다/~있다), 존댓말 금지. "
       "(2) '이 섹션은', '설명을 마무리한다', '생략한다', '제공된 정보에는' 같은 메타·상투 문구 절대 금지. "
       "(3) 이 과제는 등록 논문·특허 실적이 없으므로 특정 논문/특허 건수를 지어내지 말 것 — 대신 과제의 연구 내용과 "
       "기술 방향이 수요 해결에 어떻게 연결·활용되는지 실질적으로 서술한다. (4) 다른 머리말 없이 문단만 출력한다.")

def build_body(과제명, 설명, 수요명, 수요내용):
    u = (f"[R&D 과제]\n과제명: {과제명}\n과제설명: {설명[:700]}\n\n"
         f"[기업 기술수요]\n수요기술명: {수요명}\n수요내용: {수요내용[:400]}\n\n"
         "위를 근거로 '유사 사례 및 실적' 문단을 작성:")
    out = cm.stream_explanation([{"role": "system", "content": SYS}, {"role": "user", "content": u}],
                                max_tokens=400, temperature=0.3, top_p=0.9).strip()
    out = re.sub(r"^(유사\s*사례.*?[:：]|출력\s*[:：])\s*", "", out).strip()
    return cm.normalize_spacing(out)

def main():
    jb = json.load(open("COMPA_통합best.json", encoding="utf-8"))
    cm.load_model_blocking(progress_cb=lambda m: print(" ", m, flush=True))
    def norm(s): return re.sub(r"\s+", "", str(s or ""))
    for f, no, pid in TARGET:
        df = pd.read_pickle(f)
        idx = next(i for i, r in df.iterrows() if str(r["번호"]) == no and str(r["과제고유번호"]) == pid)
        수요명 = df.at[idx, "수요기술명"]; 기업 = df.at[idx, "기업명"]
        내용 = ""
        for e in jb.values():
            if norm(e["기업명"]) == norm(기업) and norm(e["수요기술명"]) == norm(수요명):
                내용 = e.get("수요기술 내용", ""); break
        과제명 = df.at[idx, "과제명"]; 설명 = str(df.at[idx, "과제설명문"] or "")
        new_body = build_body(과제명, 설명, 수요명, 내용)
        print("새 [유사 사례 및 실적]:\n ", new_body, flush=True)
        # 섹션 교체
        full = str(df.at[idx, "추천근거_상세"])
        parts = re.split(r"(\[[^\]]+\])", full)
        for j in range(1, len(parts), 2):
            if "유사 사례" in parts[j]:
                parts[j + 1] = " " + new_body
        newfull = cm.normalize_spacing("".join(parts))
        df.at[idx, "추천근거_상세"] = newfull
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
        print("저장:", f, flush=True)

if __name__ == "__main__":
    main()
