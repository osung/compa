# -*- coding: utf-8 -*-
"""2026 MEDITEK 수요기술 업로드 현황 XLSX → 매칭 입력용 데이터프레임 생성.

원본(2026 MEDITEK_수요기술 업로드 현황(260728).xlsx)은 COMPA_진성수요_원본.xlsx 와
컬럼 구성이 다르므로(수요기술명/사양/6T 없음, 기관 메타 위주) compa_match 가 기대하는
스키마로 정규화한다. 수요기술 내용이 비어 있는 기관은 매칭 불가로 표시만 하고 남긴다.

기업이 '확보하려는' 수요기술과 '이미 보유한' 기술·사업내용은 성격이 반대이므로
(수요를 보유역량으로 오인하면 매칭·근거가 왜곡된다) 서로 다른 파일로 분리 저장한다.
  · 수요기술   : meditek 수요기술 업로드 현황 XLSX      → MEDITEK_수요기술.pkl/.xlsx
  · 기보유기술 : meditek_firms.xlsx('기술심사 목록')     → MEDITEK_기보유기술.pkl/.xlsx

출력: MEDITEK_수요기술.pkl/.xlsx, MEDITEK_기보유기술.pkl/.xlsx
사용: python build_meditek_demands.py [--src <xlsx>] [--out-prefix MEDITEK_수요기술]
                                     [--firms <xlsx>] [--firms-out-prefix MEDITEK_기보유기술]
"""
import argparse
import os
import re

import pandas as pd

import compa_match as cm
import meditek_segment as ms

SRC = "2026 MEDITEK_수요기술 업로드 현황(260728).xlsx"
FIRMS_SRC = "meditek_firms.xlsx"          # 기업 기보유기술·사업내용(기술심사 목록)

# compa_match.match_for / gen_report 가 참조하는 수요 스키마
DEMAND_COLS = ["번호", "기업명", "수요기술명", "수요기술 내용", "수요기술 사양",
               "예상 적용 제품 및 서비스", "수요기술 분야(6T)"]
# 원본에서 보존하는 기관 메타
META_COLS = ["사업자등록번호", "대표자", "기관유형", "년도"]

# --- 기보유기술(meditek_firms.xlsx) ---
# 원본 컬럼 → 산출 컬럼(기관(업)명만 기업명으로 통일, 나머지는 원본 이름 유지)
FIRM_RENAME = {"기관(업)명": "기업명"}
FIRM_COLS = ["관리번호", "기업명", "기술유형", "기술분야", "기술명",
             "제품(기술)요약", "기술성", "제품성", "혁신성", "관련영상", "파일명"]
# '기보유기술 내용' 합본에 라벨과 함께 이어붙일 본문 섹션(원문 그대로 보존 — 요약·생성 금지)
FIRM_TEXT_COLS = ["기술명", "제품(기술)요약", "기술성", "제품성", "혁신성"]
# 값 없음을 뜻하는 자리표시 문자(관련영상 '-' 등)
_PLACEHOLDER = {"-", "--", "---", "–", "—", ".", "없음", "해당없음", "n/a", "N/A"}

# 엑셀 셀에 섞여 들어온 제어문자 잔재(form feed 등)
_CTRL = re.compile(r"_x00[0-9A-Fa-f]{2}_|[\x00-\x08\x0b\x0c\x0e-\x1f]")
# 목록 머리표(1. / 1) / - / • / ※)
_BULLET = re.compile(r"^\s*(?:\d+[.)]|[-*•·※])\s*")


def clean_text(v):
    """제어문자 제거 + 줄 단위 공백 정규화(빈 줄은 1줄로 축약). 결측은 ''."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = _CTRL.sub("", str(v))
    s = s.replace("​", "").replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in s.split("\n")]
    out, blank = [], False
    for ln in lines:
        if ln:
            out.append(ln)
            blank = False
        elif out and not blank:
            out.append("")
            blank = True
    return "\n".join(out).strip()


def derive_tech_name(content, maxlen=60):
    """수요기술 내용에서 수요기술명을 규칙 기반으로 도출 → (이름, 도출방식).

    원본에 수요기술명 컬럼이 없어 LLM 요약 대신 추출 규칙만 사용한다(임의 생성 금지).
    """
    if not content:
        return "", "없음"
    lines = [ln for ln in content.split("\n") if ln.strip()]
    if not lines:
        return "", "없음"

    # 1) 본문이 '수요기술명' 라벨을 포함하면 그 다음 줄
    for i, ln in enumerate(lines):
        if re.fullmatch(r"수요\s*기술\s*명\s*[:：]?", ln.strip()):
            if i + 1 < len(lines):
                return _trim(lines[i + 1], maxlen), "라벨"
        m = re.match(r"^수요\s*기술\s*명\s*[:：]\s*(.+)$", ln.strip())
        if m:
            return _trim(m.group(1), maxlen), "라벨"

    # 2) 대괄호 머리말 [ ... ]
    m = re.match(r"^\s*[\[【](.+?)[\]】]", lines[0])
    if m:
        return _trim(m.group(1), maxlen), "대괄호"

    # 3) 첫 유효 줄(머리표 제거)
    return _trim(_BULLET.sub("", lines[0]), maxlen), "첫줄"


def _trim(s, maxlen):
    s = _BULLET.sub("", s.strip()).strip(" .·")
    if len(s) <= maxlen:
        return s
    cut = s[:maxlen]
    sp = cut.rfind(" ")
    return (cut[:sp] if sp > maxlen * 0.6 else cut).rstrip(" ,·") + "…"


def build(src=SRC, segment="tech"):
    """원본 XLSX → 매칭 입력 DF.

    segment="tech"(기본): '주요사업분야(수요기술)' 원문 전체를 기술수요로 둔다.
    segment="llm"        : LLM으로 기술수요/회사소개/업종 분리(MEDITEK 원본에는 부적합).
    """
    raw = pd.read_excel(src)
    if segment == "llm":
        cm.load_model_blocking(progress_cb=lambda m: print("  " + m, flush=True))
    rows = []
    for r in raw.to_dict("records"):
        content = clean_text(r.get("주요사업분야(수요기술)"))
        name, how = derive_tech_name(content)
        company = clean_text(r.get("기관명"))
        if not content:
            seg = {l: [] for l in ms.LABELS}
        elif segment == "llm":
            seg = ms.segment_text(content)
            print(f"      [{r.get('No')}] {company}: "
                  + " ".join(f"{l} {len(seg[l])}문장" for l in ms.LABELS), flush=True)
        else:                                   # 기본: 원문 전체를 기술수요로
            seg = ms.segment_all_tech(content)
        rows.append({
            "번호": str(r.get("No")).strip(),
            "기업명": company,
            "수요기술명": name,
            "수요기술 내용": content,
            "수요기술 사양": "",                 # 원본에 없음
            "예상 적용 제품 및 서비스": "",         # 원본에 없음
            "수요기술 분야(6T)": "",              # 원본에 없음(보고서 단계에서 별도 부여)
            "사업자등록번호": clean_text(r.get("사업자등록번호")),
            "대표자": clean_text(r.get("대표자")),
            "기관유형": clean_text(r.get("기관유형")),
            "년도": str(r.get("년도")).strip(),
            "기업명_norm": cm.norm_name(company),  # 기업설명문 인덱스 조회 키
            "내용길이": len(content),
            "수요기술명_도출": how,
            "매칭가능": bool(content),            # 수요기술 내용 없으면 매칭 불가
            # 선정 근거 구성요소(--segment). 매칭 시 --demand-scope 로 취사선택.
            "세그_기술수요": "\n".join(seg["기술수요"]),
            "세그_회사소개": "\n".join(seg["회사소개"]),
            "세그_업종": "\n".join(seg["업종"]),
        })
    df = pd.DataFrame(rows, columns=DEMAND_COLS + META_COLS +
                      ["기업명_norm", "내용길이", "수요기술명_도출", "매칭가능",
                       "세그_기술수요", "세그_회사소개", "세그_업종"])
    df["번호"] = df["번호"].astype(str)
    return df.sort_values("번호", key=lambda s: s.astype(int)).reset_index(drop=True)


def clean_field(v):
    """clean_text + 자리표시 문자('-', '없음' 등)는 빈 값으로."""
    s = clean_text(v)
    return "" if s.strip().lower() in _PLACEHOLDER else s


def compose_firm_tech(row):
    """기보유기술 본문 섹션을 라벨과 함께 이어붙인 합본 텍스트(원문 보존)."""
    return "\n\n".join(f"[{c}]\n{row[c]}" for c in FIRM_TEXT_COLS if row[c])


def build_firms(src=FIRMS_SRC, demands=None):
    """meditek_firms.xlsx → 기업 기보유기술·사업내용 DF.

    수요기술 DF 와 같은 파일에 섞지 않고 별도 산출물로 남긴다. demands 를 주면
    기업명(법인격 표기 무시) 기준으로 대응하는 수요 '번호'를 채워 조인 키로 쓸 수 있게 한다.
    """
    raw = pd.read_excel(src).rename(columns=FIRM_RENAME)
    missing = [c for c in FIRM_COLS if c not in raw.columns]
    if missing:
        raise SystemExit(f"'{src}' 에 없는 컬럼: {missing}")

    no2 = {}
    if demands is not None:
        no2 = dict(zip(demands["기업명_norm"], demands["번호"].astype(str)))

    rows = []
    for r in raw.to_dict("records"):
        row = {c: clean_field(r.get(c)) for c in FIRM_COLS}
        content = compose_firm_tech(row)
        row["기업명_norm"] = cm.norm_name(row["기업명"])
        row["기보유기술 내용"] = content
        row["내용길이"] = len(content)
        row["기보유기술유무"] = bool(content)
        row["수요번호"] = no2.get(row["기업명_norm"], "")   # 대응 수요 없으면 ''
        rows.append(row)

    df = pd.DataFrame(rows, columns=FIRM_COLS +
                      ["기업명_norm", "기보유기술 내용", "내용길이", "기보유기술유무", "수요번호"])
    return df.sort_values("관리번호", key=lambda s: s.str.extract(r"(\d+)$")[0]
                          .astype(float)).reset_index(drop=True)


def report_firms(df, demands=None):
    ok = df[df["기보유기술유무"]]
    print(f"\n[기보유기술] 총 {len(df)}건 · 내용있음 {len(ok)}건 "
          f"· 내용없음 {len(df) - len(ok)}건")
    if len(ok):
        print(f"내용길이 min/median/max: {ok['내용길이'].min()}/"
              f"{int(ok['내용길이'].median())}/{ok['내용길이'].max()}")
    print("기술유형:", df["기술유형"].value_counts().to_dict())
    print("기술분야:", df["기술분야"].value_counts().to_dict())
    for c in FIRM_TEXT_COLS:
        n = int((df[c].str.len() > 0).sum())
        if n < len(df):
            print(f"※ '{c}' 없음 {len(df) - n}건:",
                  ", ".join(df.loc[df[c].str.len() == 0, "기업명"]))
    unmatched = df[df["수요번호"] == ""]
    if len(unmatched):
        print("※ 대응 수요기술 행 없음:",
              ", ".join(f"{n}:{c}" for n, c in zip(unmatched["관리번호"],
                                                   unmatched["기업명"])))
    if demands is not None:
        # 수요기술 내용이 비어 매칭 불가였던 기업 중 기보유기술은 확보된 경우(활용 가능성 안내)
        noc = set(demands.loc[~demands["매칭가능"], "번호"].astype(str))
        cover = df[df["기보유기술유무"] & df["수요번호"].isin(noc)]
        if len(cover):
            print("※ 수요기술 내용 없음 → 기보유기술만 확보된 기업:",
                  ", ".join(f"{n}:{c}" for n, c in zip(cover["수요번호"], cover["기업명"])))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", default=SRC)
    ap.add_argument("--out-prefix", default="MEDITEK_수요기술")
    ap.add_argument("--segment", default="tech", choices=["tech", "llm"],
                    help="tech(기본)=원문 전체를 기술수요로, llm=35B로 기술수요/회사소개/업종 분리")
    ap.add_argument("--firms", default=FIRMS_SRC,
                    help="기업 기보유기술·사업내용 XLSX(없으면 건너뜀)")
    ap.add_argument("--firms-out-prefix", default="MEDITEK_기보유기술")
    a = ap.parse_args()

    df = build(a.src, segment=a.segment)
    df.to_pickle(f"{a.out_prefix}.pkl")
    df.to_excel(f"{a.out_prefix}.xlsx", index=False)

    ok = df[df["매칭가능"]]
    print(f"총 {len(df)}건 · 매칭가능 {len(ok)}건 · 내용없음 {len(df) - len(ok)}건")
    print(f"내용길이 min/median/max: {ok['내용길이'].min()}/"
          f"{int(ok['내용길이'].median())}/{ok['내용길이'].max()}")
    print("수요기술명 도출:", df["수요기술명_도출"].value_counts().to_dict())
    if (df["내용길이"].between(1, 40)).any():
        thin = df[df["내용길이"].between(1, 40)]
        print("※ 내용 40자 미만(매칭 품질 주의):",
              ", ".join(f"{n}:{c}" for n, c in zip(thin["번호"], thin["기업명"])))
    print("※ 내용없음:", ", ".join(f"{n}:{c}" for n, c in
                                zip(df.loc[~df["매칭가능"], "번호"],
                                    df.loc[~df["매칭가능"], "기업명"])) or "-")
    if True:
        for lab in ms.LABELS:
            col = f"세그_{lab}"
            print(f"세그먼트 {lab}: 보유 {int((df[col].str.len() > 0).sum())}건 "
                  f"/ 총 {int(df[col].str.len().sum())}자")
        empty = df[df["매칭가능"] & (df["세그_기술수요"].str.len() == 0)]
        if len(empty):
            print("※ 기술수요 세그먼트 없음(--demand-scope tech 시 매칭 불가):",
                  ", ".join(f"{n}:{c}" for n, c in zip(empty["번호"], empty["기업명"])))
    print(f"저장: {os.path.abspath(a.out_prefix)}.pkl / .xlsx")

    # 기보유기술·사업내용은 수요기술과 섞지 않고 별도 파일로 저장
    if os.path.exists(a.firms):
        fdf = build_firms(a.firms, demands=df)
        fdf.to_pickle(f"{a.firms_out_prefix}.pkl")
        fdf.to_excel(f"{a.firms_out_prefix}.xlsx", index=False)
        report_firms(fdf, demands=df)
        print(f"저장: {os.path.abspath(a.firms_out_prefix)}.pkl / .xlsx")
    else:
        print(f"\n[기보유기술] '{a.firms}' 없음 → 건너뜀")


if __name__ == "__main__":
    main()
