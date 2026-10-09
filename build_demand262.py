# -*- coding: utf-8 -*-
"""2026 수요발굴지원단 기업수요 262건(xlsx 목록 + PDF 조사서) → 특허 매칭 입력용 기업정보.

xlsx 는 번호·기업명·기술번호·수요기술명만 있고, 매칭에 필요한 상세(기술 분야·수요 내용·
사양·적용 제품·도입 조건)는 PDF 조사서(수요기술 1건 = 1쪽)에만 있다. PDF 각 쪽을 파싱해
기술번호로 xlsx 와 결합한다. 본문은 원문 그대로 보존(요약·생성 없음)하고 줄바꿈만 복원한다.

출력: 수요발굴_262_기업정보.xlsx / .pkl
사용: python build_demand262.py [--xlsx <목록>] [--pdf <조사서>] [--out-prefix 수요발굴_262_기업정보]
"""
import argparse
import os
import re

import pandas as pd
import pymupdf

XLSX = "261001_수요발굴지원단 기업수요리스트_262건.xlsx"
PDF = "붙임. 2026년 수요발굴지원단 기업수요 조사서_262건.pdf"

SIX_T = ["IT", "BT", "NT", "ET", "융합"]
# 수요유형 체크 표: (항목, [(선택지 라벨, 열 위치를 잡을 대표 단어), ...]) — PDF 상 위→아래 행 순.
# 조사서 서식이 두 가지라 '수요유형' 행과 '기술지도(애로해결)' 열은 일부 쪽에만 있다.
CHOICES = [
    ("수요유형", [("기술애로해결", "기술애로해결"), ("기술도입", "기술도입"),
              ("공동연구", "공동연구"), ("기술창업", "기술창업")]),
    ("기술도입 목적", [("신제품개발", "신제품개발"), ("신공정개발", "신공정개발"),
                  ("기존제품개선", "기존제품개선"), ("기존공정개선", "기존공정개선")]),
    ("기술거래 희망 유형", [("매매(양도)", "매매(양도)"), ("전용실시", "전용실시"),
                      ("통상실시", "통상실시"), ("노하우이전", "노하우이전"),
                      ("기술지도(애로해결)", "기술지도")]),
    ("도입희망금액", [("1천만원 미만", "1천만원"), ("1천만원 이상~5천만원 미만", "~5천만원"),
                 ("5천만원 이상~1억원 미만", "~1억원"), ("1억원 이상~3억원 미만", "~3억원"),
                 ("3억원 이상", "3억원")]),
    ("도입희망시기", [("3개월 이내", "3개월"), ("6개월 이내", "6개월"), ("1년 이내", "1년"),
                 ("2년 이내", "2년"), ("3년 이내", "3년")]),
]
MARKS = ("∨", "√", "✓", "✔", "V", "v")

_HDR = re.compile(r"^\s*[\(\[]\s*((?:BT|IT|NT|ET|융합)_\d+)\s*[\)\]]", re.M)


_PARA = re.compile(r"^\s*(?:[-○〇●▪■□※•·]\s|\d+[.)]\s|[가-하][.)]\s)")


def unwrap(block):
    """PDF 줄바꿈 복원: 공백으로 끝난 줄은 다음 줄과 이어 붙이고, 아닌 줄은 문단 끝으로 본다."""
    out = ""
    for ln in block.split("\n"):
        if not ln.strip():
            if out and not out.endswith("\n"):
                out += "\n"
            continue
        # 머리표(-, ○, ※ 등)로 시작하는 줄은 앞 줄이 공백으로 끝났어도 새 문단이다
        if _PARA.match(ln) and out and not out.endswith("\n"):
            out += "\n"
        out += ln if ln.endswith(" ") else ln + "\n"
    lines = [re.sub(r"[ \t　]+", " ", x).strip() for x in out.split("\n")]
    return "\n".join(x for x in lines if x)


def between(t, start, end):
    a = t.find(start)
    if a < 0:
        return ""
    a += len(start)
    b = t.find(end, a) if end else -1
    return t[a:b if b >= 0 else None]


def checked(seg, labels):
    """'■ 라벨' 체크박스 중 선택된 라벨들."""
    out = []
    for m in re.finditer(r"■\s*([^□■\n]+)", seg):
        lab = re.sub(r"\s+", "", m.group(1))
        for L in labels:
            if lab.startswith(re.sub(r"\s+", "", L)):
                out.append(L)
                break
        else:
            out.append(m.group(1).strip())
    return out


def parse_choices(page):
    """수요유형 표의 체크(∨/√) 표시를 좌표로 읽어 행·열 선택지를 판정.

    각 행은 선택지 대표 단어들의 y 로 위치를 잡고, 그 아래~다음 행 직전 사이의 체크 표시를
    x 가 가장 가까운 선택지 열에 배정한다.
    """
    words = page.get_text("words")
    # 표 시작: '예상 적용 제품 및 서비스' 행 아래('수요유형' 라벨은 세로 가운데라 기준으로 못 씀)
    t0 = page.search_for("예상 적용")
    y_top = t0[0].y1 if t0 else 0
    words = [w for w in words if w[1] >= y_top]
    rows = []
    for key, opts in CHOICES:
        cols, used = [], set()
        for lab, anc in opts:
            cand = [w for w in words if w[4] == anc and id(w) not in used
                    and (not rows or w[1] > rows[-1][1] + 5)]
            if not cand:
                continue
            # 같은 행 안에서는 가장 위, 그다음 왼쪽 단어(금액의 '1천만원' 두 줄 중 왼쪽 열)
            y0 = min(w[1] for w in cand)
            cand = [w for w in cand if w[1] - y0 < 20]
            w = min(cand, key=lambda w: w[0])
            used.add(id(w))
            cols.append((lab, (w[0] + w[2]) / 2, w[1]))
        if len(cols) < 2:          # 이 서식에 없는 행
            continue
        rows.append((key, min(c[2] for c in cols), cols))
    marks = []
    for ch in MARKS:
        if ch in ("V", "v"):
            marks += [pymupdf.Rect(w[:4]) for w in words if w[4] == ch]
        else:
            marks += page.search_for(ch)
    res = {k: "" for k, _ in CHOICES}
    for i, (key, y, cols) in enumerate(rows):
        y_next = rows[i + 1][1] if i + 1 < len(rows) else y + 45
        sel = []
        for m in marks:
            cy = (m.y0 + m.y1) / 2
            if not (y + 4 < cy < y_next - 2):
                continue
            cx = (m.x0 + m.x1) / 2
            lab = min(cols, key=lambda c: abs(c[1] - cx))[0]
            if lab not in sel:
                sel.append(lab)
        order = [c[0] for c in cols]
        res[key] = ", ".join(sorted(sel, key=order.index))
    return res


def parse_page(page):
    t = page.get_text()
    m = _HDR.search(t)
    if not m or "수요기술 상세" not in t:
        return None
    rec = {"기술번호": m.group(1)}
    rec["수요기술명(조사서)"] = unwrap(between(t, "수요기술명\n", "수요기술 분야")).replace("\n", " ")
    six = between(t, "6T 기준", "국가")
    rec["수요기술 분야(6T)"] = ", ".join(checked(six, SIX_T))
    rec["국가과학기술표준분류(대)"] = unwrap(between(t, "대분류 기준", "중분류 기준")).lstrip(": ").replace("\n", " ")
    rec["국가과학기술표준분류(중)"] = unwrap(between(t, "중분류 기준", "전략분야")).lstrip(": ").replace("\n", " ")
    strat = between(t, "전략분야", "수요기술 상세")
    rec["전략분야"] = ", ".join(x.replace("\n", "").strip() for x in checked(strat, []))

    detail = unwrap(between(t, "수요기술 상세", "예상 적용"))
    detail = re.sub(r"^\s*내용\s*\n?", "", detail)
    rec["수요기술 상세내용"] = detail
    # 표준 서식(〇 수요기술 내용 / 〇 수요기술 사양)이면 둘로 나눠 둔다
    mm = re.search(r"[〇○o]\s*수요기술\s*내용\s*\n(.*?)(?:[〇○o]\s*수요기술\s*사양\s*\n(.*))?$", detail, re.S)
    rec["수요기술 내용"] = mm.group(1).strip() if mm else ""
    rec["수요기술 사양"] = (mm.group(2) or "").strip() if mm else ""
    m2 = re.search(r"예상\s*적용\s*제품\s*\n?\s*및\s*\n?\s*서비스(.*?)수요유형", t, re.S)
    rec["예상 적용 제품 및 서비스"] = unwrap(m2.group(1) if m2 else "").replace("\n", " ")
    rec.update(parse_choices(page))
    # 구 서식은 수요유형 선택 행 없이 '(기술도입)'으로 고정 표기
    if not rec["수요유형"] and "(기술도입)" in t:
        rec["수요유형"] = "기술도입"
    # 원문 오탈자·잘림·표기 변형 정규화(분류값만 — 본문은 손대지 않음)
    rec["전략분야"] = norm_strategy(rec["전략분야"])
    big = re.sub(r"[\s·/ㆍ‧∙]", "", rec["국가과학기술표준분류(대)"]).replace("농립", "농림")
    rec["국가과학기술표준분류(대)"] = big
    # 중분류 앞에 붙은 분류코드·구두점 잔재 제거(예: 'EH02.물관리', '.폐기물관리/자원순환')
    rec["국가과학기술표준분류(중)"] = re.sub(r"^(?:[A-Z]{2}\d{2})?\.\s*", "",
                                       rec["국가과학기술표준분류(중)"])
    return rec


STRATEGY = ["반도체∙디스플레이", "이차전지", "첨단모빌리티", "차세대원자력", "첨단바이오",
            "우주항공‧해양", "수소", "사이버 보안", "인공지능", "차세대 통신", "첨단로봇‧제조", "양자"]


def norm_strategy(v):
    """전략분야 체크값을 공식 명칭으로(서식의 '참단' 오타·'반도체∙디스' 잘림·구분자 누락 보정)."""
    key = lambda x: re.sub(r"[\s·ㆍ‧∙]", "", x).replace("참단", "첨단")
    out = []
    for part in [p for p in re.split(r",\s*", v) if p]:
        k = key(part)
        if k.startswith("기타"):
            out.append(part.strip())
            continue
        # ■ 하나에는 첫 라벨만 해당한다. 뒤에 붙은 글자는 □ 가 빠진 미선택 라벨이다
        # (예: 융합_24 '■첨단로봇제조 양자' — 양자는 미선택).
        hit = (next((s for s in STRATEGY if k.startswith(key(s))), None)
               or next((s for s in STRATEGY if key(s).startswith(k)), None))  # 잘린 라벨
        out.append(hit or part.strip())
    return ", ".join(dict.fromkeys(out))


CORP = os.path.expanduser("~/work/kipris/출원인 법인(서지)/CORP_APPLICANT.txt")
_CORP_FORM = re.compile(r"주식회사|유한책임회사|유한회사|합자회사|농업회사법인|농사회사법인|영농조합법인|"
                        r"사회적협동조합|협동조합|재단법인|사단법인|\(주\)|㈜|\(유\)|\(사\)|\(재\)")


def corp_key(name):
    """법인 형태 표기·공백·구두점을 걷어낸 상호 비교 키."""
    return re.sub(r"[\s.,·\-_()\[\]&]", "", _CORP_FORM.sub("", str(name))).upper()


def load_corp(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        hdr = f.readline().rstrip("\n").split("¶")
        for ln in f:
            p = [x.strip() for x in ln.rstrip("\n").split("¶")]
            if len(p) == len(hdr):
                rows.append(p)
    c = pd.DataFrame(rows, columns=hdr)
    c["k"] = c["출원인명"].map(corp_key)
    return c


def match_biz(names, corp):
    """기업명 → KIPRIS 출원인 법인의 사업자번호.

    상호만으로 대조하므로 같은 상호의 다른 법인이 있으면 하나로 정하지 않는다.
    법인(법인번호, 없으면 출원인코드) 기준으로 후보가 하나일 때만 사업자번호를 채우고,
    여럿이면 후보를 '사업자번호 후보'에 남긴다. KIPRIS 출원인 법인은 특허 출원 이력이 있는
    법인만 담고 있으므로, 출원 이력이 없는 기업은 찾을 수 없다.
    """
    g = corp.groupby("k")
    out = {}
    for n in names:
        k = corp_key(n)
        rec = {"사업자번호": "", "법인번호": "", "KIPRIS 출원인코드": "", "사업자번호 매칭": "미발견",
               "사업자번호 후보": "", "_kc": []}
        if k in g.groups:
            s = g.get_group(k).copy()
            s["ent"] = s["법인번호"].where(s["법인번호"] != "", s["출원인코드"])
            ents = s.groupby("ent")
            rec["_kc"] = [{"이름": x["출원인명"].iloc[0], "법인번호": x["법인번호"].iloc[0],
                           "코드": ", ".join(sorted(set(x["출원인코드"]))),
                           "biz": sorted(set(v for v in x["사업자번호"] if v))} for _, x in ents]
            if ents.ngroups == 1:
                biz = sorted(set(x for x in s["사업자번호"] if x))
                rec.update({"사업자번호": ", ".join(biz), "법인번호": s["법인번호"].iloc[0],
                            "KIPRIS 출원인코드": ", ".join(sorted(set(s["출원인코드"]))),
                            "사업자번호 매칭": "상호 단일일치" if biz else "상호 단일일치(사업자번호 없음)"})
            else:
                cand = []
                for e, x in ents:
                    biz = ",".join(sorted(set(v for v in x["사업자번호"] if v))) or "사업자번호 없음"
                    cand.append(f"{x['출원인명'].iloc[0]}[{biz}]")
                rec.update({"사업자번호 매칭": f"동명 법인 {ents.ngroups}곳", "사업자번호 후보": "; ".join(cand)})
        out[n] = rec
    return out


APOLLO_CO = os.path.expanduser("~/work/apollo/df_company_dataset_260604.pkl")
APOLLO_DESC = os.path.expanduser("~/work/apollo/company_embeddings_pro_260514_with_desc_ksic.pkl")
RERANKER = "BAAI/bge-reranker-v2-m3"
# 동명 후보 중 1위를 '추정'으로 채우는 기준: 1위/2위 적합도 비가 PICK_RATIO 이상,
# 또는 PICK_RATIO_LO 이상이면서 1위 적합도 PICK_MIN 이상. 그 밖은 후보만 남긴다.
PICK_RATIO, PICK_RATIO_LO, PICK_MIN = 10.0, 5.0, 0.25


def load_apollo(co_path, desc_path):
    co = pd.read_pickle(co_path)[["업체명", "사업자번호", "10차산업코드명", "사업목적", "보유특허명_리스트",
                                  "키워드_리스트", "지역", "설립일"]]
    de = pd.read_pickle(desc_path)[["사업자번호", "한글업체명", "기업설명문"]]
    co = co.merge(de, on="사업자번호", how="left")
    co["k"] = co["업체명"].map(corp_key)
    co["k2"] = co["한글업체명"].fillna("").map(corp_key)
    return co


def _txt(v):
    if isinstance(v, (list, tuple)) or hasattr(v, "tolist") and not isinstance(v, str):
        return ", ".join(map(str, list(v)))
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


def _profile(r):
    pats = r["보유특허명_리스트"]
    pats = list(pats)[:15] if hasattr(pats, "__len__") and not isinstance(pats, str) else []
    return (f"업종: {_txt(r['10차산업코드명'])}\n사업목적: {_txt(r['사업목적'])[:300]}\n"
            f"키워드: {_txt(r['키워드_리스트'])[:200]}\n보유특허: {'; '.join(map(str, pats))[:500]}\n"
            f"기업설명: {_txt(r['기업설명문'])[:600]}")


def fmt_biz(b):
    b = re.sub(r"\D", "", str(b))
    return f"{b[:3]}-{b[3:5]}-{b[5:]}" if len(b) == 10 else b


def resolve_biz(df, kip, co, corp):
    """KIPRIS 출원인 법인(kip)과 apollo 기업DB(co)를 합쳐 기업별 사업자번호를 정한다.

    · 두 출처가 같은 법인 하나를 가리키면 확정, 한쪽에만 하나 있으면 그 출처 단일로 채운다.
    · 같은 상호가 여럿이면 수요기술 본문과 후보 기업 프로필(업종·사업목적·키워드·보유특허·
      기업설명)을 교차 인코더로 비교해, 1위가 뚜렷할 때만 '추정'으로 채운다. 적합도는 같은
      상호 후보끼리의 상대 비교로만 쓴다(절대값은 보정되지 않아 기준으로 쓰지 않는다).
    """
    from collections import defaultdict
    idx = defaultdict(set)
    for i, (k1, k2) in enumerate(zip(co["k"], co["k2"])):
        idx[k1].add(i)
        idx[k2].add(i)
    by_biz = corp[corp["사업자번호"] != ""].groupby("사업자번호")
    ce = None
    out = {}
    for name, g in df.groupby("기업명", sort=False):
        kr = kip[name]
        kbiz = {re.sub(r"\D", "", b) for e in kr["_kc"] for b in e["biz"]}
        hits = co.loc[sorted(idx.get(corp_key(name), set()))].drop_duplicates("사업자번호")
        rec = {k: kr[k] for k in ("사업자번호", "법인번호", "KIPRIS 출원인코드")}
        rec.update({"사업자번호 근거": "", "사업자번호 후보": "", "업종": "", "지역": "", "기업설명": ""})
        scores = None
        if len(hits) > 1:
            if ce is None:
                import torch
                from sentence_transformers import CrossEncoder
                dev = "mps" if torch.backends.mps.is_available() else "cpu"
                ce = CrossEncoder(RERANKER, device=dev, max_length=1024)
                act = torch.nn.Sigmoid()
            qs = [f"{r['수요기술명']}\n{r['수요기술 상세내용'][:700]}\n적용: {r['예상 적용 제품 및 서비스']}"
                  for _, r in g.iterrows()]
            pairs = [(q, _profile(r)) for _, r in hits.iterrows() for q in qs]
            sc = ce.predict(pairs, batch_size=16, activation_fn=act).reshape(len(hits), len(qs)).max(1)
            hits = hits.assign(fit=sc).sort_values("fit", ascending=False)
            scores = hits["fit"].tolist()

        def take(row, why):
            b = fmt_biz(row["사업자번호"])
            rec.update({"사업자번호": b, "사업자번호 근거": why, "업종": _txt(row["10차산업코드명"]),
                        "지역": _txt(row["지역"]), "기업설명": _txt(row["기업설명문"])})
            if b in by_biz.groups:                       # KIPRIS 쪽 법인번호·출원인코드 보강
                x = by_biz.get_group(b)
                rec["법인번호"] = x["법인번호"].iloc[0]
                rec["KIPRIS 출원인코드"] = ", ".join(sorted(set(x["출원인코드"])))
            else:
                rec["법인번호"] = rec["KIPRIS 출원인코드"] = ""

        def cands():
            seen, out_ = set(), []
            for _, r in hits.iterrows():
                b = fmt_biz(r["사업자번호"])
                seen.add(b.replace("-", ""))
                f = f" 적합도 {r['fit']:.3f}" if "fit" in r else ""
                out_.append(f"{b} {_txt(r['10차산업코드명']) or '업종미상'}·{_txt(r['지역'])}"
                            f"{'·KIPRIS출원' if b.replace('-', '') in kbiz else ''}{f}")
            for e in kr["_kc"]:
                for b in e["biz"] or ["사업자번호 없음"]:
                    if b.replace("-", "") not in seen:
                        out_.append(f"{b} (KIPRIS 출원인 '{e['이름']}', apollo 없음)")
            return "; ".join(out_)

        n_k = len(kr["_kc"])
        if len(hits) == 0:
            if n_k == 1:
                rec["사업자번호 근거"] = "KIPRIS 단일" if kr["사업자번호"] else "KIPRIS 단일(사업자번호 없음)"
            elif n_k > 1:
                rec.update({"사업자번호": "", "법인번호": "", "KIPRIS 출원인코드": "",
                            "사업자번호 근거": f"동명 {n_k}곳 — 검토 필요", "사업자번호 후보": cands()})
            else:
                rec["사업자번호 근거"] = "미발견"
        elif len(hits) == 1:
            h = hits.iloc[0]
            if n_k == 0:
                take(h, "apollo 단일")
            elif not kbiz:                                 # KIPRIS 쪽은 사업자번호 미기재라 대조 불가
                take(h, "apollo 단일(KIPRIS 동명 법인은 사업자번호 미기재)")
            elif h["사업자번호"] in kbiz and n_k == 1:
                take(h, "KIPRIS·apollo 일치")
            elif h["사업자번호"] in kbiz:
                take(h, "apollo 단일(KIPRIS 동명 후보 중 하나)")
            else:                                          # 두 출처가 서로 다른 법인
                rec.update({"사업자번호": "", "법인번호": "", "KIPRIS 출원인코드": "",
                            "사업자번호 근거": "KIPRIS·apollo 상이 — 검토 필요", "사업자번호 후보": cands()})
        else:
            top, ratio = hits.iloc[0], scores[0] / max(scores[1], 1e-6)
            clear = ratio >= PICK_RATIO or (ratio >= PICK_RATIO_LO and scores[0] >= PICK_MIN)
            if n_k == 1 and kr["사업자번호"] and top["사업자번호"] in kbiz:
                take(top, "KIPRIS 단일·동명 중 적합도 1위")
            elif clear:
                take(top, f"동명 {len(hits)}곳 중 적합도 1위(추정)")
            else:
                rec.update({"사업자번호": "", "법인번호": "", "KIPRIS 출원인코드": "",
                            "사업자번호 근거": f"동명 {len(hits)}곳 — 검토 필요"})
            rec["사업자번호 후보"] = cands()
        out[name] = rec
    return out


MANUAL = "수요발굴_262_사업자번호_수동선택.json"   # {기업명: 사업자번호 | ""(해당 없음)}


def apply_manual(biz, path, co, corp):
    """담당자가 후보를 보고 고른 결과로 덮어쓴다. 값이 ''이면 '해당 없음'으로 비운다."""
    import json
    if not os.path.exists(path):
        return 0
    sel = json.load(open(path, encoding="utf-8"))
    co_by = co.drop_duplicates("사업자번호").set_index("사업자번호")
    by_biz = corp[corp["사업자번호"] != ""].groupby("사업자번호")
    n = 0
    for name, b in sel.items():
        if name not in biz:
            print(f"[경고] 수동선택 기업명이 목록에 없음: {name}")
            continue
        rec = biz[name]
        rec.update({"사업자번호": "", "법인번호": "", "KIPRIS 출원인코드": "", "업종": "", "지역": "", "기업설명": ""})
        if not b:
            rec["사업자번호 근거"] = "수동 확인 — 해당 없음"
        else:
            b = fmt_biz(b)
            rec["사업자번호"] = b
            rec["사업자번호 근거"] = "수동 선택"
            d = b.replace("-", "")
            if d in co_by.index:
                r = co_by.loc[d]
                rec.update({"업종": _txt(r["10차산업코드명"]), "지역": _txt(r["지역"]), "기업설명": _txt(r["기업설명문"])})
            if b in by_biz.groups:
                x = by_biz.get_group(b)
                rec["법인번호"] = x["법인번호"].iloc[0]
                rec["KIPRIS 출원인코드"] = ", ".join(sorted(set(x["출원인코드"])))
        n += 1
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default=XLSX)
    ap.add_argument("--pdf", default=PDF)
    ap.add_argument("--corp", default=CORP, help="KIPRIS 출원인 법인(서지) CORP_APPLICANT.txt")
    ap.add_argument("--apollo", default=APOLLO_CO, help="apollo 기업 데이터셋 pkl")
    ap.add_argument("--apollo-desc", default=APOLLO_DESC, help="apollo 기업설명문 pkl")
    ap.add_argument("--manual", default=MANUAL, help="사업자번호 수동 선택 JSON")
    ap.add_argument("--out-prefix", default="수요발굴_262_기업정보")
    a = ap.parse_args()

    lst = pd.read_excel(a.xlsx, header=2)
    lst = lst[pd.to_numeric(lst["번호"], errors="coerce").notna()].copy()
    lst["번호"] = lst["번호"].astype(int)
    lst["기술번호"] = lst["기술번호"].astype(str).str.strip()

    doc = pymupdf.open(a.pdf)
    recs = []
    for i, p in enumerate(doc):
        r = parse_page(p)
        if r:
            r["조사서 쪽"] = i + 1
            recs.append(r)
    pdf = pd.DataFrame(recs)
    dup = pdf["기술번호"][pdf["기술번호"].duplicated()].tolist()
    if dup:
        print("[경고] PDF 기술번호 중복:", dup)

    df = lst.merge(pdf, on="기술번호", how="left", validate="1:1")
    miss = df[df["조사서 쪽"].isna()]["기술번호"].tolist()
    extra = sorted(set(pdf["기술번호"]) - set(lst["기술번호"]))
    print(f"목록 {len(lst)}건, 조사서 {len(pdf)}쪽, 결합 누락 {len(miss)} {miss}, 목록에 없는 조사서 {extra}")

    df["조사서 쪽"] = df["조사서 쪽"].astype("Int64")

    corp = load_corp(a.corp)
    kip = match_biz(df["기업명"].unique(), corp)
    co = load_apollo(a.apollo, a.apollo_desc)
    biz = resolve_biz(df, kip, co, corp)
    print("수동 선택 반영:", apply_manual(biz, a.manual, co, corp), "곳")
    bz = pd.DataFrame.from_dict(biz, orient="index")
    df = df.join(bz, on="기업명")
    print("사업자번호 근거(기업 단위):", bz["사업자번호 근거"].str.replace(r"\d+곳", "N곳", regex=True)
          .value_counts().to_dict())

    # 원본 조사서 자체의 이상(값은 원문 그대로 두고 표시만)
    df["비고"] = ""
    body = df.set_index("기술번호")["수요기술 상세내용"]
    for k, v in body.items():
        same = [o for o, w in body.items() if o != k and w == v]
        if same:
            df.loc[df["기술번호"] == k, "비고"] += f"조사서 본문이 {','.join(same)}와 동일; "
    sim = lambda a, b: len(set(a.split()) & set(b.split())) / max(1, len(set(a.split())))
    toks = lambda x: [w for w in re.findall(r"[가-힣A-Za-z]{2,}", x)]
    for i, r in df.iterrows():
        others = df[df.index != i]
        if sim(r["수요기술명"], r["수요기술명(조사서)"]) < 0.3:
            src = others[others["수요기술명(조사서)"] == r["수요기술명(조사서)"]]["기술번호"].tolist()
            df.at[i, "비고"] += ("조사서 수요기술명 칸이 목록과 다름"
                               + (f"({','.join(src)}의 명칭이 들어간 것으로 보임 — 본문은 목록 명칭과 부합)"
                                  if src else "") + "; ")
        # 적용 제품이 자기 본문과 거의 겹치지 않고 다른 건의 값과 똑같으면 복사 오류로 본다
        # ('건강기능식품'·'S/W' 같은 짧은 일반 값은 여러 건이 같아도 정상이므로 3낱말 이상만)
        app = r["예상 적용 제품 및 서비스"]
        tk = toks(app)
        own = r["수요기술 상세내용"] + r["수요기술명"]
        # ('실증'·'시스템' 같은 일반어 한두 개만 겹치는 경우도 무관으로 본다: 겹침 20% 미만)
        if len(tk) >= 3 and sum(w in own for w in tk) / len(tk) < 0.2:
            src = others[others["예상 적용 제품 및 서비스"] == app]["기술번호"].tolist()
            if src:
                df.at[i, "비고"] += f"예상 적용 제품이 {','.join(src)}와 동일(본문과 무관 — 원본 복사 오류로 보임); "
        if r["기술번호"].split("_")[0] not in r["수요기술 분야(6T)"]:
            df.at[i, "비고"] += f"6T 체크({r['수요기술 분야(6T)']})가 기술번호 분야와 다름; "
    df["비고"] = df["비고"].str.rstrip("; ")

    cols = ["번호", "기업명", "사업자번호", "사업자번호 근거", "사업자번호 후보", "법인번호",
            "KIPRIS 출원인코드", "업종", "지역", "기업설명", "기술번호", "수요기술명", "수요기술명(조사서)", "수요기술 분야(6T)",
            "국가과학기술표준분류(대)", "국가과학기술표준분류(중)", "전략분야",
            "수요기술 내용", "수요기술 사양", "수요기술 상세내용", "예상 적용 제품 및 서비스",
            "수요유형", "기술도입 목적", "기술거래 희망 유형", "도입희망금액", "도입희망시기", "조사서 쪽", "비고"]
    df = df[cols]
    df.to_pickle(a.out_prefix + ".pkl")
    with pd.ExcelWriter(a.out_prefix + ".xlsx", engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="기업정보")
        ws = w.sheets["기업정보"]
        widths = {"번호": 6, "기업명": 22, "사업자번호": 14, "사업자번호 근거": 22, "사업자번호 후보": 50,
                  "업종": 22, "지역": 6, "기업설명": 50,
                  "법인번호": 16, "KIPRIS 출원인코드": 15, "기술번호": 10, "수요기술명": 45,
                  "수요기술명(조사서)": 45, "수요기술 분야(6T)": 10, "국가과학기술표준분류(대)": 16,
                  "국가과학기술표준분류(중)": 16, "전략분야": 16, "수요기술 내용": 60, "수요기술 사양": 60,
                  "수요기술 상세내용": 70, "예상 적용 제품 및 서비스": 30, "수요유형": 12,
                  "기술도입 목적": 14, "기술거래 희망 유형": 16, "도입희망금액": 22, "도입희망시기": 12,
                  "조사서 쪽": 8, "비고": 40}
        from openpyxl.utils import get_column_letter
        for j, c in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(j)].width = widths.get(c, 14)
        ws.freeze_panes = "C2"
    print("저장:", a.out_prefix + ".xlsx", a.out_prefix + ".pkl")


if __name__ == "__main__":
    main()
