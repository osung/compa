# -*- coding: utf-8 -*-
"""match_meditek_top10.py 산출물 검증.

산출 pkl 을 그대로 믿지 않고, 체크포인트(LLM 점수)와 코퍼스에서 순위를 독립적으로
다시 계산해 대조한다. 점수 공식도 여기서 직접 다시 구현한다(모듈 함수 재사용 금지).

  [1] 커버리지        기업 수·행 수·제외 기업
  [2] 스키마/중복     결측, 기업 내 과제 중복
  [3] 점수 정합성     최종점수·특허점수·적합도 재계산 일치
  [4] 순위 규칙       적합 우선·그룹 내 내림차순·하한
  [5] LLM 응답        누락률, 관점별 점수 분포
  [6] 특허 영향       특허 가중치 0/적합도만 순위와 반사실 비교
  [7] 표본 정성검토   기업별 상위 과제 대조

사용: python _verify_top10.py [--tag MEDITEK]
"""
import argparse
import json
import math
import os

import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer

import compa_match as cm
import match_meditek_top10 as mt

OK, BAD, WARN = "✔", "✘", "!"


def hdr(s):
    print(f"\n{'='*72}\n{s}\n{'='*72}")


def chk(cond, msg, detail=""):
    print(f"  {OK if cond else BAD} {msg}" + (f" — {detail}" if detail else ""))
    return bool(cond)


def rank_score(fit, pat, pap, prom, w, refs):
    """산출 스크립트와 독립적으로 다시 구현한 최종점수 공식."""
    pat_n = min(1.0, math.log1p(max(0, pat)) / math.log1p(refs["pat"]))
    pap_n = min(1.0, math.log1p(max(0, pap)) / math.log1p(refs["pap"]))
    prom_n = min(1.0, max(0.0, prom / 100.0))
    wsum = sum(w.values())
    total = 100.0 * (w["fit"] * fit / 100.0 + w["pat"] * pat_n
                     + w["pap"] * pap_n + w["prom"] * prom_n) / wsum
    return total, pat_n, pap_n, prom_n


def rerank(cand, w, refs, min_fit, final, title_of):
    """후보 리스트 → 선정 pid 순서. cand: [{pid, fit, pat, pap, prom, cos}]"""
    rows = []
    for x in cand:
        total = rank_score(x["fit"], x["pat"], x["pap"], x["prom"], w, refs)[0]
        rows.append(dict(x, total=total))
    rows.sort(key=lambda r: (-r["total"], -r["fit"], -r["cos"]))
    seen, sel = set(), []
    for pool in (True, False):
        for r in rows:
            if (r["fit"] >= min_fit) != pool:
                continue
            tk = title_of(r["pid"])
            if tk in seen:
                continue
            seen.add(tk)
            sel.append(r["pid"])
            if len(sel) >= final:
                break
        if len(sel) >= final:
            break
    return sel


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag", default="MEDITEK")
    ap.add_argument("--final", type=int, default=10)
    ap.add_argument("--min-fit", type=int, default=mt.MIN_FIT)
    ap.add_argument("--pat-ref", type=float, default=mt.PAT_REF,
                    help="매칭 실행 때 쓴 --pat-ref 와 같게(0=P99 자동)")
    ap.add_argument("--pap-ref", type=float, default=mt.PAP_REF)
    a = ap.parse_args()

    df = pd.read_pickle(f"{a.tag}_TOP10.pkl")
    sc = json.load(open(f"top10_{a.tag}_scores_ckpt.json", encoding="utf-8"))
    kw = json.load(open(f"top10_{a.tag}_keywords_ckpt.json", encoding="utf-8"))
    units, skipped = mt.load_units()
    units.sort(key=mt.sort_key)
    ukey = {u["기업명"]: u for u in units}
    W = {"fit": mt.W_FIT, "pat": mt.W_PAT, "pap": mt.W_PAP, "prom": mt.W_PROM}
    ok_all = []

    # ---------------------------------------------------------------- [1]
    hdr("[1] 커버리지")
    print(f"  대상 기업 {len(units)}건 · 산출 기업 {df['기업명'].nunique()}건 · 행 {len(df)}")
    ok_all.append(chk(df["기업명"].nunique() == len(units), "모든 대상 기업이 산출에 포함"))
    cnt = df.groupby("기업명").size()
    ok_all.append(chk((cnt == a.final).all(), f"기업별 {a.final}행",
                      str(cnt[cnt != a.final].to_dict() or "")))
    print(f"  제외(수요·기보유 모두 없음) {len(skipped)}건: "
          + ", ".join(f"{u['번호']}:{u['기업명']}" for u in skipped))
    print("  소스 구성:", df.drop_duplicates("기업명")["소스"].value_counts().to_dict())
    print("  키워드 개수 min/median/max:",
          f"{min(len(v) for v in kw.values())}/"
          f"{int(np.median([len(v) for v in kw.values()]))}/"
          f"{max(len(v) for v in kw.values())}")

    # ---------------------------------------------------------------- [2]
    hdr("[2] 스키마 / 중복")
    need = ["번호", "기업명", "소스", "rank", "최종점수", "적합도", "수요충족", "보유보강",
            "특허건수", "논문건수", "유망성점수", "특허점수", "과제고유번호", "과제명",
            "과제수행기관", "LLM근거", "순위구분"]
    ok_all.append(chk(all(c in df.columns for c in need), "필수 컬럼 존재",
                      str([c for c in need if c not in df.columns])))
    for c in ("과제명", "과제고유번호", "과제수행기관"):
        n = int((df[c].astype(str).str.strip() == "").sum())
        ok_all.append(chk(n == 0, f"'{c}' 공백 없음", f"{n}건"))
    dup_t = df.groupby("기업명")["과제명"].apply(lambda s: s.duplicated().sum()).sum()
    dup_p = df.groupby("기업명")["과제고유번호"].apply(lambda s: s.duplicated().sum()).sum()
    ok_all.append(chk(dup_t == 0, "기업 내 과제명 중복 없음", f"{dup_t}건"))
    ok_all.append(chk(dup_p == 0, "기업 내 과제고유번호 중복 없음", f"{dup_p}건"))
    print(f"  전체 추천 {len(df)}건 · 중복 제외 과제 {df['과제고유번호'].nunique()}개 "
          f"(여러 기업에 함께 추천된 과제 "
          f"{int((df['과제고유번호'].value_counts() > 1).sum())}개)")

    # ---------------------------------------------------------------- 코퍼스
    hdr("코퍼스 재로드(독립 재계산용)")
    c = mt.build_corpus()
    refs = mt.excellence_refs(c, a.pat_ref, a.pap_ref)
    pidx = {str(p): i for i, p in enumerate(c["pid"])}
    print("  정규화 기준:", {k: round(v, 1) for k, v in refs.items()})

    # ---------------------------------------------------------------- [3]
    hdr("[3] 점수 정합성(공식 재구현으로 재계산)")
    bad_total, bad_pat, bad_fit, bad_meta = [], [], [], []
    for r in df.to_dict("records"):
        i = pidx.get(r["과제고유번호"])
        if i is None:
            bad_meta.append(r["과제고유번호"])
            continue
        if (int(c["pat"][i]) != int(r["특허건수"])
                or int(c["pap"][i]) != int(r["논문건수"])
                or abs(float(c["promise"][i]) - float(r["유망성점수"])) > 0.02):
            bad_meta.append(r["과제고유번호"])
        t, pat_n, _, _ = rank_score(r["적합도"], r["특허건수"], r["논문건수"],
                                    r["유망성점수"], W, refs)
        if abs(t - r["최종점수"]) > 0.02:
            bad_total.append((r["기업명"], r["과제명"][:20], r["최종점수"], round(t, 2)))
        if abs(pat_n - r["특허점수"]) > 0.002:
            bad_pat.append((r["과제고유번호"], r["특허점수"], round(pat_n, 3)))
        if r["적합도"] != max(r["수요충족"], r["보유보강"]):
            bad_fit.append((r["기업명"], r["적합도"], r["수요충족"], r["보유보강"]))
    ok_all.append(chk(not bad_meta, "특허/논문/유망성이 코퍼스 값과 일치",
                      str(bad_meta[:5])))
    ok_all.append(chk(not bad_total, "최종점수 = 가중합 재계산과 일치", str(bad_total[:3])))
    ok_all.append(chk(not bad_pat, "특허점수 = log1p 정규화 재계산과 일치", str(bad_pat[:3])))
    ok_all.append(chk(not bad_fit, "적합도 = max(수요충족, 보유보강)", str(bad_fit[:3])))

    # 소스별 관점 점수 정합
    d_only = df[df["소스"] == "수요"]
    h_only = df[df["소스"] == "기보유"]
    ok_all.append(chk((d_only["보유보강"] == 0).all(), "수요만 있는 기업 → 보유보강 0"))
    ok_all.append(chk((h_only["수요충족"] == 0).all(), "기보유만 있는 기업 → 수요충족 0"))

    # ---------------------------------------------------------------- [4]
    hdr("[4] 순위 규칙")
    bad_order, bad_group, bad_gate = [], [], []
    for name, g in df.groupby("기업명", sort=False):
        g = g.sort_values("rank")
        if list(g["rank"]) != list(range(1, len(g) + 1)):
            bad_order.append(f"{name}(rank 불연속)")
        grp = list(g["순위구분"])
        if "적합" in grp and "보충" in grp and grp.index("보충") < len(grp) - grp.count("보충"):
            bad_group.append(name)
        for key in ("적합", "보충"):
            s = g[g["순위구분"] == key]["최종점수"].tolist()
            if s != sorted(s, reverse=True):
                bad_order.append(f"{name}({key} 내 정렬)")
        if (g[g["순위구분"] == "적합"]["적합도"] < a.min_fit).any():
            bad_gate.append(name)
    ok_all.append(chk(not bad_order, "rank 연속 + 그룹 내 최종점수 내림차순",
                      str(bad_order[:3])))
    ok_all.append(chk(not bad_group, "'적합'이 '보충'보다 앞", str(bad_group[:3])))
    ok_all.append(chk(not bad_gate, f"'적합' 행은 모두 적합도 ≥ {a.min_fit}",
                      str(bad_gate[:3])))
    fill = df[df["순위구분"] == "보충"]
    print(f"  보충 선정 {len(fill)}행 / {len(df)}행 "
          f"({fill['기업명'].nunique()}개 기업: "
          + ", ".join(f"{n}({k}건)" for n, k in fill.groupby('기업명').size().items()) + ")")

    # ---------------------------------------------------------------- [5]
    hdr("[5] LLM 응답 / 점수 분포")
    tot = sum(len(v) for v in sc.values())
    miss = sum(1 for v in sc.values() for s in v.values() if s[2] == "(LLM 응답 누락)")
    zero = sum(1 for v in sc.values() for s in v.values() if s[0] == 0 and s[1] == 0)
    ok_all.append(chk(miss / max(tot, 1) < 0.02, "LLM 응답 누락률 < 2%",
                      f"{miss}/{tot} ({miss/max(tot,1)*100:.1f}%)"))
    print(f"  채점 후보 {tot}건 · 양쪽 0점 {zero}건({zero/max(tot,1)*100:.0f}%) · "
          f"근거 빈 문자열 {sum(1 for v in sc.values() for s in v.values() if not s[2])}건")
    fits = np.array([max(s[0], s[1]) for v in sc.values() for s in v.values()])
    print(f"  후보 적합도 분포: 평균 {fits.mean():.1f} · "
          f"≥90 {int((fits >= 90).sum())} · 70-89 {int(((fits >= 70) & (fits < 90)).sum())} · "
          f"40-69 {int(((fits >= 40) & (fits < 70)).sum())} · <40 {int((fits < 40).sum())}")
    both = df[df["소스"] == "수요+기보유"]
    if len(both):
        print(f"  수요+기보유 기업 {both['기업명'].nunique()}개의 TOP{a.final}: "
              f"수요충족 우세 {int((both['수요충족'] > both['보유보강']).sum())}행 · "
              f"보유보강 우세 {int((both['보유보강'] > both['수요충족']).sum())}행 · "
              f"동점 {int((both['보유보강'] == both['수요충족']).sum())}행")
        print("   → 기보유기술이 실제로 매칭 근거로 작동했는지 확인용"
              "(보유보강 우세 행이 있으면 수요만으로는 뽑히지 않았을 과제)")

    # ---------------------------------------------------------------- [6]
    hdr("[6] 산출 재현 + 특허 가중치 반사실(counterfactual)")
    print("  · pro-sroberta 로드(코사인 재계산)…")
    model = SentenceTransformer(cm.MODEL_DIR, device="cpu")
    model.eval()
    encode = cm.make_encoder(model)

    def title_of(pid):
        return cm.title_key(c["pname"][pidx[str(pid)]], False)

    repro, cf_pat0, cf_fit = [], [], []
    rows_cf = []
    for u in units:
        key, name = u["키"], u["기업명"]
        got = sc.get(key, {})
        q = encode(cm.SEP.join(kw[key]))
        cand = []
        for pid, (d, h, _r) in got.items():
            i = pidx.get(str(pid))
            if i is None:
                continue
            cand.append({"pid": str(pid), "fit": max(d, h), "pat": int(c["pat"][i]),
                         "pap": int(c["pap"][i]), "prom": float(c["promise"][i]),
                         "cos": float(np.dot(c["M"][i], q))})
        got_pids = df[df["기업명"] == name].sort_values("rank")["과제고유번호"].tolist()
        mine = rerank(cand, W, refs, a.min_fit, a.final, title_of)
        repro.append((name, mine == got_pids, len(set(mine) & set(got_pids))))
        # 반사실 ①: 특허 가중치 0 / ②: 적합도만
        w0 = dict(W, pat=0.0)
        p0 = rerank(cand, w0, refs, a.min_fit, a.final, title_of)
        pf = rerank(cand, {"fit": 1.0, "pat": 0.0, "pap": 0.0, "prom": 0.0},
                    refs, a.min_fit, a.final, title_of)
        cf_pat0.append(len(set(got_pids) - set(p0)))
        cf_fit.append(len(set(got_pids) - set(pf)))
        pats = [x["pat"] for x in cand if x["pid"] in got_pids]
        allp = [x["pat"] for x in cand]
        rows_cf.append({"기업명": name, "소스": u["소스"],
                        "재현": "일치" if mine == got_pids else "불일치",
                        "동일선정": len(set(mine) & set(got_pids)),
                        "특허0_교체": cf_pat0[-1], "적합도만_교체": cf_fit[-1],
                        "TOP평균특허": round(float(np.mean(pats)), 1),
                        "후보평균특허": round(float(np.mean(allp)), 1)})
    cf = pd.DataFrame(rows_cf)
    ok_all.append(chk(all(x[1] for x in repro),
                      "산출 TOP 순서를 독립 재계산으로 100% 재현",
                      ", ".join(f"{n}({k}/{a.final} 동일)" for n, o, k in repro if not o)))
    print(cf.to_string(index=False))
    print(f"\n  특허 가중치 0 으로 두면 평균 {np.mean(cf_pat0):.1f}/{a.final}건이 교체됨 "
          f"(특허가 순위에 실제로 기여) · 적합도만 쓰면 평균 {np.mean(cf_fit):.1f}/{a.final}건 교체")
    print(f"  TOP{a.final} 평균 특허건수 {cf['TOP평균특허'].mean():.1f} vs "
          f"후보 200 평균 {cf['후보평균특허'].mean():.1f} vs "
          f"코퍼스 평균 {c['pat'].mean():.1f}")
    ok_all.append(chk(cf["TOP평균특허"].mean() > cf["후보평균특허"].mean(),
                      "TOP 선정 과제의 평균 특허건수가 후보 평균보다 높음"))

    # ---------------------------------------------------------------- [7]
    hdr("[7] 표본 정성검토 — 기업별 상위 3건")
    for name, g in df.groupby("기업명", sort=False):
        u = ukey[name]
        g = g.sort_values("rank").head(3)
        head = f"■ [{u['번호'] or u['관리번호']}] {name} ({u['소스']})"
        print(f"\n{head}")
        if u["수요기술명"]:
            print(f"   수요: {u['수요기술명'][:70]}")
        if u["기보유기술명"]:
            print(f"   보유: {u['기보유기술명'][:70]}")
        print(f"   키워드: {', '.join(kw[u['키']][:10])}")
        for r in g.to_dict("records"):
            print(f"   {r['rank']}. [{r['최종점수']:.1f} | 적합 {r['적합도']}"
                  f"(수요 {r['수요충족']}/보유 {r['보유보강']}) | 특허 {r['특허건수']}건] "
                  f"{r['과제명'][:52]}")
            print(f"      근거: {r['LLM근거'][:88]}")

    hdr("검증 요약")
    print(f"  통과 {sum(1 for x in ok_all if x)}/{len(ok_all)} 항목")
    print("  " + (f"{OK} 모든 검증 항목 통과" if all(ok_all)
                  else f"{BAD} 실패 항목 있음 — 위 {BAD} 표시 확인"))


if __name__ == "__main__":
    main()
