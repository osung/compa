# -*- coding: utf-8 -*-
"""상세 매칭 근거의 메타/상투 문구(생략한다/이 섹션은/서술할 수 없다/제공된 정보에는…) 섹션을
자연스러운 반말 문장으로 재생성·교체. 논문/특허 있으면 인용, 없으면 기술 연결성으로 서술(날조 금지)."""
import re, os, json, time, pickle
import pandas as pd
import compa_match as cm

S = os.environ.get("COMPA_SCRATCH", "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad")
CK = f"{S}/meta_sections_ckpt.json"
PKLS = ["COMPA_필터10_최종추천.pkl", "COMPA_필터11_20_최종추천.pkl", "COMPA_필터21_78_최종추천.pkl"]
META = re.compile(r"이 섹션은|설명을 마무리|서술을 마무리|생략(한다|합니다)|서술할 수 없|설명할 수 없|분석할 수 없|제공된 정보에는")

def norm(s): return re.sub(r"\s+", "", str(s or ""))
jb = json.load(open("COMPA_통합best.json", encoding="utf-8"))
dem_meta = {(norm(e["기업명"]), norm(e["수요기술명"])): (e.get("수요기술 내용", ""), e.get("수요기술 사양", ""))
            for e in jb.values()}
pmeta = pickle.load(open("project_match_data_260612.pkl", "rb"))

SYS = ("너는 기술이전 보고서의 상세 매칭 근거 문단을 쓰는 한국어 AI다. 주어진 섹션의 목적에 맞춰, 과제와 기업 수요를 "
       "근거로 2~3문장을 자연스럽게 쓴다. 규칙: (1) 반드시 반말 한다체(~다/~한다/~이다/~된다/~있다), 존댓말 금지. "
       "(2) '이 섹션은', '설명을 마무리한다', '생략한다', '제공된 정보에는', '서술할 수 없다' 같은 메타·상투·회피 문구 절대 금지. "
       "(3) 없는 논문/특허 실적을 지어내지 말 것. 실적이 주어지면 그 성과와 수요의 연결을, 없으면 과제의 연구 내용·기술 방향이 "
       "수요 해결에 어떻게 활용·연결되는지 실질적으로 서술한다. (4) 머리말 없이 문단만 출력한다.")

PURPOSE = {
    "연관성": "기업의 사업·역량과 과제 목표·내용의 기술적 연관성",
    "수요기술 사양 적합성": "수요기술 사양(수치·기준)을 과제가 충족·근접하는 정도",
    "추천 과제의 우수성": "과제의 유망성·연구성과·수행기관 역량 등 강점",
    "유사 사례 및 실적": "과제의 논문·특허 등 실적(또는 연구 내용)이 수요 해결에 주는 시사점",
}

def gen(sec, 과제명, 설명, 수요명, 내용, 사양, npaper, npatent, paps, pats):
    실적 = f"논문 {npaper}건, 특허 {npatent}건"
    if paps: 실적 += " / 논문예: " + "; ".join(paps[:2])
    if pats: 실적 += " / 특허예: " + "; ".join(pats[:2])
    u = (f"[섹션] {sec} — {PURPOSE.get(sec,'')}\n"
         f"[과제] 과제명: {과제명}\n설명: {설명[:600]}\n연구성과: {실적}\n\n"
         f"[기업 수요] 수요기술명: {수요명}\n내용: {내용[:300]}\n사양: {사양[:200]}\n\n"
         f"'{sec}' 문단 작성:")
    for temp in (0.3, 0.6):
        o = cm.normalize_spacing(cm.stream_explanation(
            [{"role": "system", "content": SYS}, {"role": "user", "content": u}],
            max_tokens=420, temperature=temp, top_p=0.9).strip())
        o = re.sub(r"^\[?%s\]?\s*[:：]?\s*" % re.escape(sec), "", o).strip()
        if o and not META.search(o) and not re.search(r"(?<!아)니다", o):
            return o
    return None

def main():
    ck = json.load(open(CK)) if os.path.exists(CK) else {}
    frames = {f: pd.read_pickle(f) for f in PKLS}
    jobs = []
    for f, df in frames.items():
        for i, r in df.iterrows():
            parts = re.split(r"(\[[^\]]+\])", str(r["추천근거_상세"] or ""))
            for j in range(1, len(parts), 2):
                sec = parts[j].strip("[]")
                body = parts[j + 1] if j + 1 < len(parts) else ""
                if META.search(body):
                    jobs.append((f, i, j, sec))
    print(f"메타 섹션 대상: {len(jobs)} (체크포인트 {len(ck)})", flush=True)
    cm.load_model_blocking(progress_cb=lambda m: print("  " + m, flush=True))
    n = 0
    for f, i, j, sec in jobs:
        n += 1
        df = frames[f]; r = df.loc[i]
        pid = str(r["과제고유번호"]); key = f"{r['번호']}::{pid}::{sec}"
        if key in ck:
            continue
        내용, 사양 = dem_meta.get((norm(r["기업명"]), norm(r["수요기술명"])), ("", ""))
        pm = pmeta.get(pid, {})
        new = gen(sec, r["과제명"], str(r["과제설명문"] or ""), r["수요기술명"], 내용, 사양,
                  pm.get("논문건수", 0) or 0, pm.get("특허건수", 0) or 0,
                  pm.get("논문명_리스트") or [], pm.get("특허명_리스트") or [])
        ck[key] = new or ""
        print(f"[{n}/{len(jobs)}] {key} {'OK' if new else 'FAIL(유지)'}", flush=True)
        if n % 5 == 0:
            json.dump(ck, open(CK, "w"), ensure_ascii=False)
    json.dump(ck, open(CK, "w"), ensure_ascii=False)

    # 반영: 섹션 교체
    for f, df in frames.items():
        for i, r in df.iterrows():
            parts = re.split(r"(\[[^\]]+\])", str(r["추천근거_상세"] or ""))
            changed = False
            for j in range(1, len(parts), 2):
                sec = parts[j].strip("[]"); pid = str(r["과제고유번호"])
                key = f"{r['번호']}::{pid}::{sec}"
                if key in ck and ck[key]:
                    parts[j + 1] = " " + ck[key]; changed = True
            if changed:
                df.at[i, "추천근거_상세"] = cm.normalize_spacing("".join(parts))
        df.to_pickle(f); df.to_excel(f.replace(".pkl", ".xlsx"), index=False)
        print("저장", f, flush=True)
    left = sum(1 for f, df in frames.items() for _, r in df.iterrows() if META.search(str(r["추천근거_상세"] or "")))
    print("완료. 메타 문구 잔존 행:", left, flush=True)

if __name__ == "__main__":
    main()
