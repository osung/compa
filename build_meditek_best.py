# -*- coding: utf-8 -*-
"""MEDITEK 매칭 결과(pkl) → 보고서 입력 JSON(MEDITEK_통합best.json).

_build_full_inputs.py 의 MEDITEK 판. 6T(demand_field.json)는 생성하지 않는다
(보고서를 분야별 장으로 나누지 않음).

여러 tag 를 겹쳐 쓸 수 있다. 뒤에 오는 pkl 이 같은 번호를 덮어쓴다.
  예) 기본 all scope 위에 9·17·30 만 tech scope 결과로 교체
      python build_meditek_best.py --pkl COMPA_MEDITEK_최종추천.pkl \
                                   --pkl COMPA_MEDITEK_tech_최종추천.pkl
"""
import argparse
import json
import os

import pandas as pd

DEMANDS_PKL = "MEDITEK_수요기술.pkl"
OUT = "MEDITEK_통합best.json"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pkl", action="append", default=None,
                    help="최종추천 pkl(여러 번 지정 가능, 뒤가 앞을 덮어씀)")
    ap.add_argument("--demands", default=DEMANDS_PKL)
    ap.add_argument("--out", default=OUT)
    a = ap.parse_args()
    pkls = a.pkl or ["COMPA_MEDITEK_최종추천.pkl"]

    dm = pd.read_pickle(a.demands)
    dm["번호"] = dm["번호"].astype(str)
    dmeta = dm.set_index("번호").to_dict("index")

    best, source = {}, {}
    for f in pkls:
        if not os.path.exists(f):
            raise SystemExit(f"없는 파일: {f}")
        df = pd.read_pickle(f)
        df["번호"] = df["번호"].astype(str)
        df["과제고유번호"] = df["과제고유번호"].astype(str)
        for no in df["번호"].unique():
            g = df[df["번호"] == no].sort_values("rank")
            meta = dmeta.get(no, {})
            best[no] = {
                "기업명": meta.get("기업명", g.iloc[0]["기업명"]),
                "수요기술명": meta.get("수요기술명", g.iloc[0]["수요기술명"]),
                # 매칭에 실제 투입된 수요 텍스트(scope 반영분)를 그대로 싣는다
                "수요기술 내용": str(g.iloc[0].get("수요기술 내용", "") or ""),
                "수요기술 사양": "",
                "top5": [{"rank": int(r["rank"]), "과제고유번호": str(r["과제고유번호"]),
                          "과제명": r["과제명"], "수행기관": r["과제수행기관"],
                          "LLM점수": int(r["LLM점수"]), "판단근거": r["LLM판단근거"],
                          "과제설명문": r["과제설명문"], "추천근거_상세": r["추천근거_상세"]}
                         for r in g.to_dict("records")],
            }
            source[no] = f

    if os.path.exists(a.out):
        os.replace(a.out, a.out + ".bak")
    with open(a.out, "w", encoding="utf-8") as fp:
        json.dump(best, fp, ensure_ascii=False, indent=1)

    n_rec = sum(len(v["top5"]) for v in best.values())
    n_proj = len({t["과제고유번호"] for v in best.values() for t in v["top5"]})
    print(f"{a.out}: 수요 {len(best)}건 · 추천 {n_rec}건 · 중복제외 과제 {n_proj}개")
    for f in pkls:
        ks = sorted((k for k, v in source.items() if v == f), key=int)
        print(f"  {f}: {ks}")


if __name__ == "__main__":
    main()
