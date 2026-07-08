# -*- coding: utf-8 -*-
"""45건 표'매칭 근거' 열용 한 문장(긍정형 매칭 근거)을 few-shot 프롬프트로 35B 생성.
상세 4섹션 근거(추천근거_상세)는 그대로 유지하고, 표 셀용 짧은 근거만 별도 생성."""
import json, os, time
import compa_match as cm

SCRATCH = "/private/tmp/claude-501/-Users-osung-work-compa/d6ed121c-12e4-45b4-b2fb-535b7554627c/scratchpad"
OUT = f"{SCRATCH}/regen_out.json"

def log(*a): print(*a, flush=True)

tgt = {f"{t['번호']}::{t['pid']}": t for t in json.load(open(f"{SCRATCH}/regen_targets.json"))["regen"]}
done = json.load(open(OUT))

SYS = (
    "너는 기업 기술수요와 국가 R&D 과제의 '매칭 근거'를 한 문장으로 요약하는 한국어 AI다. "
    "이 과제가 왜 해당 수요에 추천되는지, 두 대상이 공유하는 핵심 기술·목적을 근거로 긍정적으로 서술한다.\n"
    "작성 규칙:\n"
    "1) 60자 이내 한 문장. 명사형/음슴체 종결(예: '~ 기술이 수요와 부합', '~에 활용 가능', '~ 요구사항과 일치').\n"
    "2) 과제 개요·수행기관·기간·분야 등 '설명'은 절대 쓰지 말고, 수요와 과제의 '접점(매칭 이유)'만 쓴다.\n"
    "3) 부정 평가('~ 부재', '~ 미포함', '~ 불일치', '~ 미흡', '다름')는 쓰지 않는다. "
    "완전히 일치하지 않는 경우에도 공유하는 기술적 접점을 중심으로 '~에 활용 가능', '~ 기반 마련', "
    "'~ 부분 부합' 처럼 기여 가능성으로 표현한다.\n"
    "4) 문장 하나만 출력. 따옴표·머리기호·부연 금지.\n"
    "5) few-shot 예시는 형식·톤 참고용이다. 예시의 고유명사·문구를 출력에 재사용하지 말고, "
    "반드시 현재 입력된 수요·과제 내용만 근거로 작성한다."
)

def demo(dem, proj, ans):
    u = f"[기업 기술수요]\n{dem}\n\n[R&D 과제]\n{proj}\n\n매칭 근거(한 문장):"
    return [{"role": "user", "content": u}, {"role": "assistant", "content": ans}]

FEWSHOT = (
    demo("수요기술명: 도라지 사포닌 정제·표준화 및 분말 제형화 기술\n수요기술 내용: 다년근 도라지 유래 플라티코딘 D 등 사포닌을 표준화하고 분말 제형으로 안정화",
         "과제명: 도라지 사포닌 추출 및 정제 공정 표준화 연구\n과제설명: 도라지에서 사포닌을 고효율로 추출·정제하고 분말화하는 표준 공정 개발",
         "도라지 사포닌 추출·정제·분말화 표준 공정이 수요 요구사항과 직접 일치")
    + demo("수요기술명: 엣지 컴퓨터 기반 차량 번호판 인식 기술\n수요기술 내용: 저전력 엣지 디바이스에서 실시간 번호판 OCR을 위한 딥러닝 모델 경량화",
           "과제명: 엣지 인공지능 자동화 기술\n과제설명: 신경망 모델 압축·양자화 및 자동 신경망 설계로 엣지 디바이스 저전력 실시간 추론 구현",
           "모델 압축·양자화 기반 엣지 저전력 실시간 추론 기술이 경량 OCR 수요와 부합")
    + demo("수요기술명: 견과류 코팅형 인지기능 개선 소재화 기술\n수요기술 내용: 인지기능 개선 기능성 성분을 견과류에 코팅해 고령친화 제품으로 소재화",
           "과제명: 감탄닌 소재 인지기능 개선 연구\n과제설명: 감탄닌 천연물의 인지기능 개선 효능을 규명하고 기능성 소재로 개발",
           "인지기능 개선 천연소재 연구가 수요의 기능성 소재 확보에 활용 가능")
)

cm.load_model_blocking(progress_cb=lambda m: log("  " + m))

for n, (k, v) in enumerate(done.items(), 1):
    if v.get("매칭근거_short"):
        log(f"[{n}/{len(done)}] skip"); continue
    t = tgt.get(k, {})
    dem = (f"수요기술명: {v['수요기술명']}\n"
           f"수요기술 내용: {t.get('수요기술 내용','')[:600]}\n"
           f"수요기술 사양: {t.get('수요기술 사양','')[:350]}")
    proj = f"과제명: {v['과제명']}\n과제설명: {t.get('과제설명문_fallback','')[:600]}"
    user = f"[기업 기술수요]\n{dem}\n\n[R&D 과제]\n{proj}\n\n매칭 근거(한 문장):"
    msgs = [{"role": "system", "content": SYS}] + FEWSHOT + [{"role": "user", "content": user}]
    out = cm.stream_explanation(msgs, max_tokens=120, temperature=0.0, top_p=1.0).strip()
    out = out.strip().strip('"').strip("'").split("\n")[0].strip()
    if out.startswith("매칭 근거"):
        out = out.split(":", 1)[-1].strip()
    v["매칭근거_short"] = out[:120]
    json.dump(done, open(OUT, "w"), ensure_ascii=False, indent=1)
    log(f"[{n}/{len(done)}] {k} :: {out}")

log("SHORT REASON DONE", len(done))
