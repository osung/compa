# -*- coding: utf-8 -*-
"""COMPA 최종 Top5 보고서(docx) 통합 생성 스크립트.

편집 전 원본 보고서(COMPA_최종Top5_보고서_원본.docx)에 우리가 정한 모든 서식·내용 규칙을
순서대로 적용해 최종 보고서를 재현한다(전체 390건). 매칭 근거는 기존 산출물을 재사용하며,
--regen 옵션 시 Qwen3.5-35B-A3B(MLX)로 45건 근거를 새로 생성한 뒤 반영한다.

파이프라인(각 단계는 검증된 모듈 스크립트, 파일→파일):
  [--regen] _regen_run.py → _regen_shortreason.py   (45건 근거 재생성; 기본은 재사용)
  1 _patch_docx.py     : 72건 근거 채움 + 표헤더 '매칭 이유'→'매칭 근거'
  2 _strip_source.py   : '출처' 표기 전면 삭제
  3 _add_disclaimer.py : 제목 아래 AI 유의사항 박스
  4 _add_pagenum.py    : 요약/제1장부터 새 구역·페이지번호(1부터)
  5 _demand_table.py   : 수요별 기업명/내용/사양 2열 표
  6 _add_year_col.py   : Top5 표에 과제수행년도 열
  7 _fmt_top5.py       : Top5 정렬 + 수행년도 2줄
  8 _justify.py        : 상세근거·수요내용/사양·매칭근거 양쪽맞춤
  9 _del_summary.py    : '전체 요약' 섹션 삭제  (+ 제목 변경)
  10 _add_proj_table.py: 각 Top 과제 7필드 정보표

사용:
  python build_report.py                 # 기존 근거 재사용(모델 불필요)
  python build_report.py --regen         # Qwen3.5로 45건 근거 재생성 후 빌드
  python build_report.py --out 파일.docx  # 출력 경로 지정
"""
import os, sys, shutil, runpy

HERE = os.path.dirname(os.path.abspath(__file__))
SCRATCH = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"

BASE = os.path.join(HERE, "COMPA_최종Top5_보고서_원본.docx")
NEW_TITLE = "COMPA 매칭데이 기술수요조사 최종 매칭 보고서"

def arg(flag, default=None):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv else default

REGEN = "--regen" in sys.argv
OUT = arg("--out", os.path.join(HERE, "COMPA_최종Top5_보고서.docx"))

STEPS = [
    "_patch_docx.py", "_strip_source.py", "_add_disclaimer.py", "_add_pagenum.py",
    "_demand_table.py", "_add_year_col.py", "_fmt_top5.py", "_justify.py",
    "_del_summary.py", "_add_proj_table.py",
]

def run_module(script, argv):
    print(f"\n===== 실행: {script} {argv} =====", flush=True)
    sys.argv = [script] + argv
    runpy.run_path(os.path.join(HERE, script), run_name="__main__")

def set_title(path):
    from docx import Document
    d = Document(path)
    p = d.paragraphs[0]
    runs = p.runs
    tgt = max(runs, key=lambda r: len(r.text)) if any(r.text for r in runs) else runs[-1]
    for r in runs:
        r.text = NEW_TITLE if r is tgt else ""
    d.save(path)
    print(f"  제목 변경 → {NEW_TITLE}")

def main():
    assert os.path.exists(BASE), f"원본 입력 없음: {BASE}"

    if REGEN:
        print("### --regen: Qwen3.5-35B-A3B 로 45건 근거 재생성 ###", flush=True)
        run_module("_regen_run.py", [])
        run_module("_regen_shortreason.py", [])

    cur = os.path.join(SCRATCH, "_pipe_00.docx")
    shutil.copy(BASE, cur)
    for i, step in enumerate(STEPS, 1):
        nxt = os.path.join(SCRATCH, f"_pipe_{i:02d}.docx")
        run_module(step, [cur, nxt])
        if step == "_del_summary.py":       # 요약 삭제 직후 제목 변경(과거 순서와 동일)
            set_title(nxt)
        cur = nxt

    shutil.copy(cur, OUT)
    print(f"\n✔ 최종 보고서 생성 완료 → {OUT}")

if __name__ == "__main__":
    main()
