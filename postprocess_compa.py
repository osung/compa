# -*- coding: utf-8 -*-
"""COMPA 최종추천 xlsx 표준 산출물 생성.

담당자별로 다음 두 파일을 만든다(원본 최종추천 xlsx 는 보존):
  - COMPA_{담당자}_최종추천_보완.xlsx        : 정해진 15컬럼 순서 + (보완)추천근거_상세
  - COMPA_{담당자}_최종추천_보완_기업별.xlsx : 기업별 탭 분리 + 문자열 멀티라인(wrap)

컬럼 규칙
  - '추천근거_상세_보완' 컬럼이 있으면 기존 '추천근거_상세' 를 버리고 그것을 '추천근거_상세'로 사용.
  - 아래 ORDER 순서로 재정렬(존재하는 컬럼만).
"""
import os
import sys
import pandas as pd
import split_by_company as sbc

ORDER = ["번호", "기업명", "기술번호", "수요기술명", "키워드", "과제고유번호", "과제명",
         "과제수행기관", "LLM점수", "LLM판단근거", "추천근거_상세", "과제설명문",
         "유사도_과제코사인", "유사도_기업코사인", "유망성점수"]


def clean_flat(src, dst):
    """src xlsx → 15컬럼 정리본(dst). 보완 근거 컬럼이 있으면 승격."""
    df = pd.read_excel(src)
    if "추천근거_상세_보완" in df.columns:
        if "추천근거_상세" in df.columns:
            df = df.drop(columns=["추천근거_상세"])
        df = df.rename(columns={"추천근거_상세_보완": "추천근거_상세"})
    cols = [c for c in ORDER if c in df.columns]
    df = df[cols]
    df.to_excel(dst, index=False)
    return df


def make_deliverables(assignee, src=None):
    """담당자 최종추천 → 정리본(_보완) + 기업별 탭본(_보완_기업별) 생성.
    src 를 명시하면 그 파일을 원본으로 사용(파이프라인이 방금 쓴 최종추천 전달용).
    미지정 시, 이미 정리된 _보완.xlsx 가 있으면 그것을(멱등), 없으면
    파이프라인 산출물 COMPA_{a}_최종추천.xlsx 를 원본으로 삼는다."""
    flat = f"COMPA_{assignee}_최종추천_보완.xlsx"
    tabs = f"COMPA_{assignee}_최종추천_보완_기업별.xlsx"
    raw = f"COMPA_{assignee}_최종추천.xlsx"
    if src is None:
        src = flat if os.path.exists(flat) else raw
    if not os.path.exists(src):
        raise FileNotFoundError(f"원본 없음: {src}")
    clean_flat(src, flat)
    sbc.split_file(flat, tabs)
    return flat, tabs


def main():
    assignees = (sys.argv[1].split(",") if len(sys.argv) >= 2
                 else ["김민주", "이중연"])
    for a in assignees:
        a = a.strip()
        print(f"[후처리] {a}")
        make_deliverables(a)
    print("완료.")


if __name__ == "__main__":
    main()
