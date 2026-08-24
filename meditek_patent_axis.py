# -*- coding: utf-8 -*-
"""과제 특허 임베딩 기반 매칭 점수 축.

목적: **질의와 유사하고 권리가 살아 있는 특허를 많이 가진 과제**가 높은 점수를 받게
한다. 특허 '건수'만 세는 기존 필터(1건 이상)는 거절·취하된 특허 1건짜리 과제도
통과시키는데, 그런 과제는 실제로 기술이전 협의가 불가능하다.

소스: ../apollo/nice_patent_260813.pkl
  보고서 특허 표가 쓰는 nice_patent_overview_rnd_linked_abstract_promise.pkl 과
  행 수가 같은(403,024) 같은 계보의 파일에 norm_embed(768d L2)·키워드·특허설명이
  더 붙어 있다. 코퍼스 1,803 과제 전체를 덮는다(100%, 특허 4,509건, 손상 0건).

축 구성 — 품질을 **곱하기**로 넣는다(합 1.0 인 세 계수의 의미는 그대로):
    qual_i = 0.55 + 0.30·법적상태_i + 0.15·최신성_i        ∈ [0.55, 1.00]
      0.55 관련성 기본값 · 0.30 권리 프리미엄 · 0.15 최신성 프리미엄
      법적상태_score = 등록 1.00 / 소멸 0.55 / 공개 0.45 / 거절 0.10 / 취하·포기 0.05
      → 등록 특허 1건은 취하 특허 1건의 약 1.42배로 세어진다(최신성 평균 0.75 기준)

  더하기(0.55·유사강도 + 0.30·권리 + 0.15·최신성)로 만들면 안 된다. 실측 결과
  권리·최신성은 LLM 적합도와 상관이 없거나 음(−0.037/−0.032)인데 — 기업 적합성이
  아니라 자산 품질을 재는 값이라 당연하다 — 가중의 45%를 차지하면서 관련성 신호를
  절반으로 깎았다(적합도 순위상관 +0.129 → +0.074, 유사특허 개수 단조성 +0.947 → +0.624).
  곱하기로 바꾸면 관련성을 전혀 잃지 않고(+0.127) 개수 단조성(+0.947)과 권리 변별을
  동시에 지킨다. 자세한 실측표는 docs/매칭개선_특허임베딩축_20260822.md 참고.

※ 원본의 `유망성점수`·`모과제유망성` 은 쓰지 않는다. 합성 유망성점수에는
  모과제유망성(= 그 특허가 나온 과제의 유망성)이 섞여 있어서, 이미 별도 축으로
  0.15 가중을 주는 과제 유망성이 두 번 반영된다(상관 +0.405). 이중 계산을 피하려고
  구성요소인 법적상태·최신성만 골라 쓴다.

유사도 에너지 — 질의 q, 과제 p 의 특허 i:
    τ_q   = 코퍼스 전체 특허 유사도의 P90           (질의별 적응 임계값)
    rel_i = max(sim_i − τ_q, 0)                    (임계 초과분만 인정)
    E_p   = Σ_i rel_i · qual_i        ← 유사하고 · 유망하고 · 많을수록 커진다
    e_n   = 1 − exp(−E_p / E_ref),  E_ref = E>0 인 과제의 P95
      → 단조증가하며 상한에 붙지 않는다. 개수 효과가 상위에서도 계속 살아 있게 하려고
        P95 에서 자르지 않고 포화 곡선을 쓴다(E=E_ref 에서 0.632, 2배 0.865).
    pat_n = 0.75·e_n + 0.25·mx_n·qual_best
      mx_n = (max sim − τ)/(1 − τ) — 특허가 1건뿐인 과제(코퍼스 중위값이 1건)가
      에너지만으로 과소평가되지 않게 최고 유사도 항을 섞고, 거기에도 그 특허의
      품질 계수를 곱한다.

τ·E_ref 를 질의마다 다시 잡는 이유: 짧은 키워드 질의는 코퍼스 중심 근처에 놓여
모든 특허와 높은 유사도를 갖는다(실측 P90 이 질의에 따라 0.38~0.98). 절대 임계값은
쓸 수 없고, 질의별 백분위로만 비교 가능한 척도가 나온다.
"""
import os
import re

import numpy as np
import pandas as pd

PAT_EMB = os.environ.get(
    "COMPA_PATENT_EMB", "/Users/osung/work/apollo/nice_patent_260813.pkl")

W_PSIM, W_PLEGAL, W_PRECENT = 0.55, 0.30, 0.15   # 품질 계수 qual = 0.55+0.30·권리+0.15·최신성
TAU_PCT = 90.0       # 질의별 관련 임계값 백분위
EREF_PCT = 95.0      # 에너지 정규화 기준 백분위
MAX_MIX = 0.25       # 에너지 안에서 최고 유사도 항의 비중
TOP_K = 5            # 과제당 에너지에 기여하는 특허 수 상한(유사도 상위 K건)
                     # 코퍼스 과제당 특허는 중위 1건인데 최대 153건짜리 우산형 과제가
                     # 있다. 상한이 없으면 산하 연구실 특허를 대량 누적한 인재양성·
                     # 기관지원 사업이 유리해진다. 실측상 K 를 두어도 적합도 순위상관은
                     # +0.133 로 변하지 않아(K=2~무제한 전부 동일) 신호 손실이 없다.
LIVE_MIN = 0.40      # 권리 생존 하한 — 등록 1.00 · 공개 0.45 는 통과.
                     # ALIVE_STATES 로 이미 걸러지므로 사후 확인용 안전장치다.
DIM = 768

# 절대 읽지 않는 컬럼 — 과제 유망성 축과 겹친다(이중 계산 방지)
BANNED = ("유망성점수", "유망성점수_기본", "유망성_raw", "모과제유망성")
USE = ("과제번호", "특허출원번호", "특허명", "norm_embed", "법적상태_score",
       "최신성_score", "특허등록상태명")
# 축에 넣는 특허 상태 — 매칭 코퍼스 필터와 같은 규칙(출원 계류 또는 등록).
# 권리가 없는 특허는 기술이전 대상이 아니므로 유사도 에너지에 기여해서는 안 된다.
ALIVE_STATES = ("등록", "공개")


def _pids(v):
    if isinstance(v, (list, tuple, set, np.ndarray)):
        return [str(x).strip() for x in v if str(x).strip()]
    return re.findall(r"\d{6,}", str(v))


class PatentAxis:
    """코퍼스 과제 순서(pid_list)에 맞춘 특허 임베딩 집계기."""

    def __init__(self, P, pidx, legal, recent, n_proj, n_pat_used, n_bad, states,
                 titles=None, states_of=None):
        self.P, self.pidx = P, pidx
        self.legal, self.recent = legal, recent
        # 품질 계수: 관련성 기본 0.55 + 권리 0.30 + 최신성 0.15 → [0.55, 1.00]
        self.qual = W_PSIM + W_PLEGAL * legal + W_PRECENT * recent
        self.n_proj, self.n_pat = n_proj, n_pat_used
        self.n_bad, self.states = n_bad, states
        # 근거문에 인용할 특허명 — 질의와 실제로 유사한 특허만 골라 프롬프트에 넣는다
        self.titles = np.asarray(titles if titles is not None else [""] * len(pidx))
        self.states_of = np.asarray(states_of if states_of is not None
                                    else [""] * len(pidx))
        self.has = np.bincount(pidx, minlength=n_proj) > 0
        # 과제 단위 권리 생존 여부(질의와 무관) — 코퍼스 필터가 '출원 1건 이상'이라
        # 거절·취하된 특허만 가진 과제도 들어온다. 보고서에서 "이전 협의 가능"으로
        # 읽히면 안 되므로 최고 권리 상태를 따로 남겨 검증에서 주의로 표시한다.
        self.legal_max = np.zeros(n_proj)
        np.maximum.at(self.legal_max, pidx, legal)

    def score(self, q):
        """질의 임베딩(768d L2) → (pat_n, 진단 dict). 둘 다 길이 n_proj."""
        q = np.asarray(q, dtype=np.float32).ravel()
        sims = np.clip(self.P @ q, -1.0, 1.0)
        tau = float(np.percentile(sims, TAU_PCT))
        rel = np.maximum(sims - tau, 0.0).astype(np.float64)

        n = self.n_proj
        # 과제별 관련도 상위 TOP_K 건만 에너지에 기여시킨다
        o = np.lexsort((-rel, self.pidx))          # 과제 오름차순 · 관련도 내림차순
        grp = self.pidx[o]
        head = np.searchsorted(grp, np.arange(n), side="left")
        sel = o[(np.arange(len(o)) - head[grp]) < TOP_K]
        ps_, rl_ = self.pidx[sel], rel[sel]
        E = np.bincount(ps_, weights=rl_ * self.qual[sel], minlength=n)    # 품질 가중
        E0 = np.bincount(ps_, weights=rl_, minlength=n)                    # 진단용(무가중)
        SL = np.bincount(ps_, weights=rl_ * self.legal[sel], minlength=n)
        SR = np.bincount(ps_, weights=rl_ * self.recent[sel], minlength=n)
        cnt = np.bincount(ps_, weights=(rl_ > 0).astype(np.float64), minlength=n)

        # 과제별 최고 유사도 특허 — 관련 특허가 없는 과제의 대체 경로
        mx = np.full(n, -1.0)
        np.maximum.at(mx, self.pidx, sims.astype(np.float64))
        best = sims >= mx[self.pidx] - 1e-9
        bq = np.zeros(n)
        bl = np.zeros(n)
        br = np.zeros(n)
        bq[self.pidx[best]] = self.qual[best]
        bl[self.pidx[best]] = self.legal[best]
        br[self.pidx[best]] = self.recent[best]

        pos = E > 0
        e_ref = float(np.percentile(E[pos], EREF_PCT)) if pos.any() else 1.0
        e_n = 1.0 - np.exp(-E / max(e_ref, 1e-9))          # 포화하지만 상한에 안 붙는다
        mx_n = np.clip((mx - tau) / max(1.0 - tau, 1e-9), 0.0, 1.0)

        pat_n = (1.0 - MAX_MIX) * e_n + MAX_MIX * mx_n * bq
        pat_n = np.where(self.has, np.clip(pat_n, 0.0, 1.0), 0.0)
        # 관련 특허가 있으면 그 특허들의 권리·최신성, 없으면 최고 유사도 특허의 값
        legal_n = np.where(pos, np.divide(SL, np.where(pos, E0, 1.0)), bl)
        recent_n = np.where(pos, np.divide(SR, np.where(pos, E0, 1.0)), br)
        return pat_n.astype(np.float64), {
            "tau": tau, "e_ref": e_ref, "E": E, "E0": E0, "cnt": cnt, "e_n": e_n,
            "legal": np.clip(legal_n, 0, 1), "recent": np.clip(recent_n, 0, 1),
            "max": mx, "mx_n": mx_n, "qual_best": bq,
            "legal_max": self.legal_max,
        }


    def relevant(self, q, i, n=3, min_sim=None):
        """질의 q 와 유사한 과제 i 의 특허 상위 n건 → [(특허명, 상태, 유사도)…].

        임계값(τ, 질의별 P90)을 넘는 특허만 돌려준다. 넘는 특허가 없으면 빈 목록이다 —
        근거문에 억지로 특허를 끼워 넣지 않기 위해서다.
        """
        q = np.asarray(q, dtype=np.float32).ravel()
        sims = np.clip(self.P @ q, -1.0, 1.0)
        tau = float(np.percentile(sims, TAU_PCT)) if min_sim is None else float(min_sim)
        sel = np.flatnonzero((self.pidx == i) & (sims > tau))
        if not len(sel):
            return []
        sel = sel[np.argsort(-sims[sel])][:n]
        out, seen = [], set()
        for j in sel:
            t = str(self.titles[j]).strip()
            if not t or t in seen:
                continue
            seen.add(t)
            out.append((t, str(self.states_of[j]), float(sims[j])))
        return out


def load(pid_list, src=PAT_EMB, log=print, members=None):
    """코퍼스 과제 목록 → PatentAxis. pid_list 의 순서가 축 인덱스가 된다.

    members: {대표 과제고유번호: [연차 과제고유번호…]} — 다년차 통합 시 전 연차의
    특허를 대표 과제에 모아 붙인다. 주지 않으면 과제고유번호 1:1 로만 연결한다.
    """
    if not os.path.exists(src):
        raise SystemExit(f"[중단] 특허 임베딩 파일 없음: {src}")
    pos = {}
    for i, p in enumerate(pid_list):
        for q in (members or {}).get(str(p), [str(p)]):
            pos[str(q)] = i
    log(f"· 특허 임베딩 로드… ({os.path.basename(src)})"
        + (f" · 연차 통합 {len(pos)}개 과제번호 → {len(pid_list)}개 대표"
           if members else ""))
    df = pd.read_pickle(src)
    banned = [c for c in BANNED if c in df.columns]
    miss = [c for c in USE if c not in df.columns]
    if miss:
        raise SystemExit(f"[중단] 특허 임베딩에 없는 컬럼: {miss}")
    df = df[list(USE)]                       # 유망성 계열은 아예 들고 오지 않는다
    log(f"  전체 {len(df)}건 · 배제한 유망성 컬럼 {len(banned)}개: {banned}")

    embs, idx, lg, rc, bad, states, seen = [], [], [], [], 0, {}, set()
    ttl, sof = [], []
    n_dead = 0
    for t in df.itertuples(index=False):
        hit = [pos[p] for p in _pids(t.과제번호) if p in pos]
        if not hit:
            continue
        e = np.asarray(t.norm_embed, dtype=np.float32).ravel()
        if e.size != DIM:
            bad += 1
            continue
        key = str(t.특허출원번호).strip()
        st = str(t.특허등록상태명).strip()
        if st not in ALIVE_STATES:      # 거절·취하·포기·소멸은 축에서 제외
            n_dead += 1
            continue
        for j in hit:
            if key and (j, key) in seen:      # 같은 과제 안 중복 출원번호 제거
                continue
            seen.add((j, key))
            embs.append(e)
            idx.append(j)
            lg.append(float(t.법적상태_score))
            rc.append(float(t.최신성_score))
            ttl.append(str(t.특허명).strip())
            sof.append(st)
            states[st] = states.get(st, 0) + 1
    del df
    if not embs:
        raise SystemExit("[중단] 코퍼스 과제에 연결된 특허 임베딩이 0건")

    P = np.vstack(embs)
    nrm = np.linalg.norm(P, axis=1)
    off = int((np.abs(nrm - 1.0) > 1e-3).sum())
    if off:                                   # L2 정규화가 안 된 행은 직접 정규화
        P = P / np.where(nrm[:, None] > 0, nrm[:, None], 1.0)
        log(f"  (참고) L2 정규화되지 않은 {off}건 재정규화")
    ax = PatentAxis(P, np.asarray(idx, dtype=np.int64),
                    np.asarray(lg), np.asarray(rc), len(pid_list), len(embs), bad,
                    states, titles=ttl, states_of=sof)
    log(f"  코퍼스 특허 {ax.n_pat}건 · 과제 {int(ax.has.sum())}/{len(pid_list)} "
        f"({ax.has.mean()*100:.1f}%) · 손상 임베딩 {bad}건 · "
        f"상태 미달로 제외 {n_dead}건(거절·취하·포기·소멸)")
    log(f"  등록상태: " + " · ".join(f"{k} {v}" for k, v in
                                  sorted(states.items(), key=lambda kv: -kv[1])[:6]))
    dead = int(((ax.legal_max < LIVE_MIN) & ax.has).sum())
    log(f"  권리 생존 과제 {int(((ax.legal_max >= LIVE_MIN) & ax.has).sum())}"
        f"/{int(ax.has.sum())} · 거절·취하뿐인 과제 {dead}건")
    log(f"  품질계수 qual = {W_PSIM} + {W_PLEGAL}·권리 + {W_PRECENT}·최신성 "
        f"→ 범위 {ax.qual.min():.3f}~{ax.qual.max():.3f} 평균 {ax.qual.mean():.3f}")
    log(f"  τ=P{TAU_PCT:g} · E_ref=P{EREF_PCT:g} · 최고유사도 혼합 {MAX_MIX} "
        f"· 과제당 기여 특허 상한 {TOP_K}건")
    return ax
