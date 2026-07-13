# COMPA 최종 Top5 보고서(docx) 작성 규칙

`COMPA_최종Top5_보고서.docx` 를 생성·편집할 때 적용한 규칙 정리.
기업 진성수요 78건 × 수요당 추천 Top5(총 390건) + 근거를 담은 보고서.
편집은 원본 재생성이 아니라 **python-docx 로 in-place 패치**(정상 항목의 문단/셀을 복제해 서식 유지)로 수행한다.
전면 재생성은 `gen_report.py`(docx) / `gen_report_pdf.py`(PDF)로 하며, 둘은 동일 데이터·동일 출판물 디자인을 공유한다.

---

# ▶ 실행 가이드 (자기완결 런북)

**다른 컴퓨터에서 이 저장소(스크립트) + 아래 외부 데이터/모델을 준비하면, Qwen LLM 매칭 → docx·PDF 보고서까지 전부 재현된다.** (파이프라인 세부 규칙은 §0~§10.)

전체 흐름: `진성수요 xlsx` → **①매칭(Qwen 35B)** → `COMPA_통합best.json`(정본) → **②입력 캐시** → **③보고서(docx·PDF)**.

## A. 환경 설정

- **HW/OS**: Apple Silicon(권장, MLX 백엔드) 또는 CUDA GPU(vLLM). RAM ≥ 32GB(피크 ~30GB), 디스크 ~15GB(데이터+모델).
- **Python** 3.10+.
- **패키지**: `pip install -r requirements.txt` (numpy·pandas·openpyxl·sentence-transformers·scikit-learn + python-docx·reportlab·fonttools·pillow·pypdf) **그리고 LLM 백엔드 하나**: Apple `pip install mlx-lm` / CUDA `pip install "vllm>=0.19.0"`.
- **Qwen LLM**: 기본 `mlx-community/Qwen3.5-35B-A3B-4bit`(MLX). 첫 실행 시 HuggingFace에서 자동 다운로드(네트워크 필요, 캐시 `~/.cache/huggingface`). CUDA는 `Qwen/Qwen3.5-35B-A3B-GPTQ-Int4`(vLLM). 교체는 환경변수 `MV_MLX_MODEL`/`MV_VLLM_MODEL`.
- **임베딩 모델**: 저장소의 `pro-sroberta/` 디렉토리(SentenceTransformers). 재랭커 `BAAI/bge-reranker-v2-m3`는 첫 사용 시 자동 다운로드.
- **스크래치 경로**(임시 캐시·폰트): `export COMPA_SCRATCH="$PWD/_scratch"; mkdir -p "$COMPA_SCRATCH/fonts"`.
- **apollo 데이터 경로**: `export COMPA_APOLLO_DIR=/경로/apollo` (아래 B의 apollo 파일 위치).
- **폰트(보고서 디자인, Noto Sans/Serif KR)** — VF 내려받아 static Regular/Bold 인스턴싱:
  ```bash
  cd "$COMPA_SCRATCH/fonts"
  curl -L -o NotoSansKR-VF.ttf  "https://github.com/google/fonts/raw/main/ofl/notosanskr/NotoSansKR%5Bwght%5D.ttf"
  curl -L -o NotoSerifKR-VF.ttf "https://github.com/google/fonts/raw/main/ofl/notoserifkr/NotoSerifKR%5Bwght%5D.ttf"
  python - <<'PY'
  from fontTools.varLib.instancer import instantiateVariableFont
  from fontTools.ttLib import TTFont
  for vf, base in [("NotoSansKR-VF.ttf","NotoSansKR"), ("NotoSerifKR-VF.ttf","NotoSerifKR")]:
      for w, n in {400:"Regular", 700:"Bold"}.items():
          f = TTFont(vf); instantiateVariableFont(f, {"wght": w}, inplace=True); f.save(f"{base}-{n}.ttf")
  PY
  ```
  (docx를 Word에서 열 땐 시스템에 Noto Sans/Serif KR 설치 권장. LibreOffice는 검증용, 선택.)

## B. 필요 데이터 파일 (외부·대용량, git 제외 — 별도 확보 필요)

**매칭 엔진용**(저장소 루트 = `COMPA_DATA_DIR`, 기본 현재 디렉토리):

| 파일 | 내용 |
|---|---|
| `COMPA_진성수요_원본.xlsx` | 입력 수요 78건(담당자·수요기술명·내용·사양·6T 분야) |
| `pro-sroberta/` | 임베딩 모델 디렉토리 |
| `public_RnD_embeddings_pro_with_desc_260708.pkl` | 과제 임베딩+과제명+제출년도+키워드+유망성 + **연구수행주체·연구개발단계**(필터·정보표) |
| `company_embeddings_pro_260514_with_desc.pkl` | 기업 설명문(상세근거용) |
| `project_match_data_260612.pkl` | 과제 메타(수행기관·논문/특허·상위비율) |

**보고서 상세필드·특허·연구책임자용**(`$COMPA_APOLLO_DIR`):

| 파일 | 내용 |
|---|---|
| `df_project_dataset_260602.pkl` | 총연구비·연구시작/종료일·표준분류(중)·수행기관 |
| `df_project_all_bizno_260310.pkl` | 과학기술표준분류(대) |
| `public_RnD_PI_260610.pkl` | 연구책임자명·국가연구자번호 |
| `df_pr_patent_260710_detail.pkl` | 특허 출원/등록 상세 |
| `APOLLO 로고.png` | 로고(저장소 루트) |

## C. 스크립트 (저장소 포함)

`compa_match.py`(매칭 엔진: 키워드추출·합집합검색·재랭킹·35B 채점·근거생성), `rematch_filtered.py`(필터 재매칭 드라이버), `_build_full_inputs.py`·`_patent_prep.py`(보고서 입력 캐시 생성), `gen_report.py`(docx), `gen_report_pdf.py`(PDF), `_rebuild_final.py`·`_fix_projtable_pids.py`(정본 동기화/정정). 경로는 모두 `COMPA_SCRATCH`·`COMPA_APOLLO_DIR` 환경변수를 따른다.

## D. 실행 절차 (end-to-end)

**① 매칭 (필터 재매칭)** — 78수요를 배치로. 수요 1건당 ~5–20분(35B), 항목별 체크포인트로 중단·재개 가능:
```bash
python rematch_filtered.py --start 1  --n 10 --tag 필터10
python rematch_filtered.py --start 11 --n 10 --tag 필터11_20
python rematch_filtered.py --start 21 --n 58 --tag 필터21_78
```
(필터 조건 = §0. `rematch_filtered.py` 상단 `ALLOW`(연구수행주체)·`YEAR_MIN`(2020) 확인. 태그 3개는 `_build_full_inputs.py`의 `BATCHES` 목록과 일치해야 함.)

**② 보고서 입력 통합**:
```bash
python _build_full_inputs.py   # → COMPA_통합best.json(정본) + $COMPA_SCRATCH/{pid_fields,demand_field}.json
python _patent_prep.py         # → $COMPA_SCRATCH/pid_patents.json
```

**③ 로고 crop**(PDF가 참조하는 `apollo_top.png`; docx를 먼저 생성하면 `make_top_logo()`가 자동 생성하므로 생략 가능):
```bash
python - <<'PY'
import os, numpy as np
from PIL import Image
S = os.environ["COMPA_SCRATCH"]; im = Image.open("APOLLO 로고.png").convert("RGBA"); a = np.array(im)[:, :, 3]
rows = np.where((a > 16).sum(1) > 5)[0]; sp = np.where(np.diff(rows) > 20)[0]
band = rows[:sp[0] + 1] if len(sp) else rows; y0, y1 = int(band[0]), int(band[-1])
cols = np.where((a[y0:y1 + 1] > 16).sum(0) > 2)[0]; x0, x1 = int(cols[0]), int(cols[-1]); p = 12
im.crop((max(0, x0 - p), max(0, y0 - p), min(im.width, x1 + p), min(im.height, y1 + p))).save(f"{S}/apollo_top.png")
PY
```

**④ 보고서 생성**:
```bash
COMPA_REPORT_OUT=COMPA_최종보고서.docx python gen_report.py       # docx
COMPA_PDF_OUT=COMPA_최종보고서.pdf     python gen_report_pdf.py   # PDF(배포용)
```

**⑤ (선택) 검증** — LibreOffice로 docx→PDF 변환 후 페이지 렌더 확인(§9-7):
```bash
soffice --headless --convert-to pdf --outdir "$COMPA_SCRATCH" COMPA_최종보고서.docx
```

## D-1. 보고서 생성 입력 JSON (4종) — 명세와 생성 방법

`gen_report.py`(docx)·`gen_report_pdf.py`(PDF)는 **아래 JSON 4종만** 읽어 보고서를 만든다(LLM·apollo pkl 직접 참조 없음; 폰트·로고 정적 자산은 별도). apollo pkl은 이 JSON을 **만들 때만**(②) 쓰인다.

| JSON | 위치 | 생성 스크립트 | 내용 | 소스 |
|---|---|---|---|---|
| `COMPA_통합best.json` | 저장소 루트 | **①매칭** 후 `_build_full_inputs.py`(또는 `_rebuild_final.py`) | 수요 78건 × Top5(과제·수행기관·rank·출처) + **매칭 근거(판단근거)·상세 4섹션(추천근거_상세)·과제설명문** = 매칭·근거의 **정본** | 매칭 산출물 `COMPA_필터*_최종추천.pkl` |
| `pid_fields.json` | `$COMPA_SCRATCH` | `_build_full_inputs.py` | pid→ `연구개발단계·연구수행주체·연구책임자명·국가연구자번호` | apollo `public_RnD_embeddings_pro_with_desc_260708.pkl`, `public_RnD_PI_260610.pkl` |
| `pid_patents.json` | `$COMPA_SCRATCH` | `_patent_prep.py` | pid→ 특허 실적 목록(등록 우선·출원 병기, 다년도 전 연도) | apollo `df_pr_patent_260710_detail.pkl` |
| `demand_field.json` | `$COMPA_SCRATCH` | `_build_full_inputs.py` | 수요번호→ 6T 분야(BT/IT/NT/ET/융합) | `COMPA_진성수요_원본.xlsx` |

생성(= 런북 ②단계):
```bash
python _build_full_inputs.py   # COMPA_통합best.json + $COMPA_SCRATCH/{pid_fields,demand_field}.json
python _patent_prep.py         # $COMPA_SCRATCH/pid_patents.json
```
- `COMPA_통합best.json`은 **① 매칭(Qwen 35B)** 산출물을 통합한 것 = 매칭+근거의 정본. `_build_full_inputs.py`의 `BATCHES`(예: `COMPA_필터10/필터11_20/필터21_78_최종추천.pkl`)를 실제 태그에 맞춰 지정.
- 나머지 3종은 apollo pkl·입력 xlsx에서 pid/수요번호로 뽑은 **캐시**라 매칭 없이도 갱신 가능(단 apollo 데이터 필요).
- 스키마 예시 — `COMPA_통합best.json`:
  ```json
  { "1": { "기업명": "...", "수요기술명": "...", "수요기술 내용": "...", "수요기술 사양": "...",
           "top5": [ { "rank": 1, "과제고유번호": "1345086170", "과제명": "...", "수행기관": "...",
                       "LLM점수": 90, "출처": "기존+신규", "판단근거": "...(매칭 근거 한 줄)",
                       "과제설명문": "...", "추천근거_상세": "[연관성] ... [수요기술 사양 적합성] ..." } ] } }
  ```
  `pid_fields.json`: `{ "1345086170": {"연구개발단계":"기초연구","연구수행주체":"대학","연구책임자명":"홍길동","국가연구자번호":"1000..."} }`
  `pid_patents.json`: `{ "1345086170": [ {"특허명":"...","기관":"...","국가":"KR","출원번호":"...","출원일":"2021.03","등록번호":"...","등록일":"2022.05","상태":"등록"} ] }`
  `demand_field.json`: `{ "1": "BT", "2": "BT", "17": "IT", ... }`

## E. 참고

- **보고서만 재생성**: 이미 `COMPA_통합best.json`이 있으면 ①(매칭)을 건너뛰고 ②~④만 실행하면 된다(단, apollo 데이터·폰트·로고는 필요). apollo 원본조차 없더라도 위 **JSON 4종이 이미 있으면 ③(보고서 생성)만으로 재현**된다.
- 대용량 데이터·모델·산출물(pkl/xlsx/docx/pdf/폰트)은 **git 제외**. 다른 컴퓨터에선 B의 데이터를 별도로 확보해야 매칭이 재현된다.
- 요청 모델이 Qwen3.6-35B-A3B였으나 미보유/네트워크 이슈로 본 세션은 3.5로 실행. 모델 교체는 `MV_MLX_MODEL`.

---

## 0. 매칭 대상 필터 (재매칭 시 필수)

재매칭(`rematch_filtered.py`)은 후보 코퍼스를 아래 조건으로 걸러 35B가 필터 통과 과제만 채점·선정하게 한다.
- **최근 5년**: `제출년도 ≥ 2020` (오늘 기준 최근 5년; `public_RnD_embeddings_pro_with_desc_260708.pkl` 의 `제출년도`).
- **연구수행주체**: `{대학, 출연연구소, 국공립연구소, 정부부처}` 만 (같은 파일 `연구수행주체`). 중소기업·중견기업·대기업·기타 제외.
- 필터 통과 약 228k건(전체 96만 중). 재매칭은 수요 1건당 ~5–20분(35B, top200 채점+근거). 산출물 `COMPA_필터*_최종추천.pkl` → `_build_full_inputs.py` 로 `COMPA_통합best.json`(정본) 통합.
- 산출물은 항목별 체크포인트(`compa2stage_*_ckpt.json`)로 중단·재개 가능.

## 1. 문서 구조 / 구역

- **구역 1 (표지·목차)**: 제목 → 유의사항 박스 → 개요 문단 → 목차. **페이지 번호 없음.**
- **구역 2 (본문)**: `■ 전체 요약` 부터 문서 끝(분야별 장 + 수요기술 상세). `■ 전체 요약` 앞에 구역 나누기(nextPage) 삽입.
- 계층: `Heading 1`=분야(장)·전체 요약, `Heading 2`=`[수요 N] 수요기술명`, `Heading 3`=`TopN. 과제명`.
- 각 수요 블록 순서: **[수요 N] 제목 → 수요 정보 표 → `■ 매칭 관점` → `■ 최종 추천 과제 Top5` 표 → TopN 상세 블록(×5)**.
- **페이지 나누기**(§1-1 참조): 각 `[수요 N]`(Heading 2)과 각 `TopN` 상세 블록(Heading 3)은 **새 페이지에서 시작**한다. 즉 수요 시작 페이지에는 수요 정보 표+Top5 표만, 이후 각 추천 과제 상세는 **1과제 = 1페이지**.

### 1-1. 페이지 나누기 / 1페이지 제한 (조판)

- **새 수요기술은 새 페이지에서 시작**: `[수요 N]` 제목(Heading 2) 문단에 `pageBreakBefore`(python-docx `paragraph_format.page_break_before = True`). 앞 문단 `space_before` 는 0 으로.
- **매칭 과제 상세(Top1~5)는 각각 새 페이지에서 시작**: `TopN. 과제명`(Heading 3) 문단에 동일하게 `pageBreakBefore`. Top5 표 다음 페이지부터 Top1 이 시작되어, `Top1~5 설명`이 새 페이지에서 열린다.
- **한 과제 상세는 1페이지를 넘기지 않는다**: 상세 블록(정보표 7행 + `적합성 판단` + `상세 매칭 근거` 4섹션)이 최장 근거(≈2,300자)에서도 한 페이지에 들어가도록 아래 조판값을 적용.
  - 상세 블록 본문 폰트 **9pt**(정보표·적합성·4섹션 공통), 4섹션 **줄간격 1.12**, 섹션 간 간격 **3pt**.
  - 문서 여백 **상·하 0.7in / 좌·우 0.9in** 로 축소해 세로 공간 확보(표 최대폭 8640dxa 는 그대로 페이지 내 수렴).
  - 위 값은 최장 블록 렌더 확인으로 검증됨. 근거가 더 길어져 넘칠 경우 폰트를 8.5pt 로, 줄간격을 1.1 로 추가 축소.
- **검증**: docx→PDF 변환 후, 한 수요 내 연속한 `TopN`·`Top(N+1)` 제목의 시작 페이지 차이가 **정확히 1** 이면 과제별 1페이지 제한 충족(2 이상이면 오버플로우). 모든 Heading 2/3 에 `page_break_before=True` 인지도 확인.

## 2. 페이지 번호

- 구역 2 `sectPr` 에 `pgNumType w:start="1"` → **전체 요약이 1페이지**.
- 구역 2 하단(footer) **가운데** 에 `PAGE` 필드. footer 는 이전 구역과 링크 해제(`is_linked_to_previous=False`).
- 구역 1 footer 는 비움(표지·목차 무번호).

## 3. 첫 페이지 유의사항(경고문)

- 위치: 제목 문단 **바로 아래**. 중복 삽입 방지(문구 존재 시 skip).
- 박스: 4면 단선 테두리 색 `C00000`(sz 12), 음영 fill `FCE9E9`, 문단 전후 간격 120.
- 텍스트: 1줄차 `※ 유의사항`(굵게, `C00000`, 10.5pt) + 줄바꿈 + 본문(`7F1D1D`, 9.5pt).
- 본문 문구:
  > 본 보고서는 APOLLO 인공지능을 이용해서 생성한 보고서로 사실과 다르거나 오류가 있을 수 있습니다. 참고용으로만 활용하시고, 정확한 정보는 관련 자료를 통해 확인하시기 바랍니다. 본 보고서는 AI 생성 내용의 정확성을 보증하지 않으며, 이를 근거로 한 판단 의사결정의 책임은 이용자에게 있습니다.

## 4. 매칭 근거 / 상세 추천 근거 (핵심)

- **모든 근거 텍스트는 로컬 35B(MLX `Qwen3.5-35B-A3B-4bit`, `compa_match.py`)가 생성**한다. **작성자(사람/Claude)가 직접 쓰지 않는다.**
- **표의 '매칭 근거'(한 줄)**: 긍정형 매칭 근거. 부정 평가(`~부재`, `~미포함`, `~불일치`, `~미흡`, `다름`) 금지 → 부분 부합도 `~에 활용 가능`, `~ 기반 마련`, `~ 부분 부합` 등 **공유 접점·기여 가능성**으로 서술. 60자 이내 한 문장, 명사형/음슴체 종결. **few-shot 프롬프트**로 톤·형식 고정(예시 문구 재사용 금지).
  - 배경: 단건 `llm_score`(“표면 일치는 저점”)는 저점수 후보에 비판적 한 줄을 내므로 '추천' 표엔 부적합. 그래서 별도 긍정형 생성 패스로 교체.
  - **구현(강제)**: `compa_match.gen_match_reason(demand, proj)` — 긍정 few-shot으로 한 문장 생성. `match_for` 가 이 함수로 `LLM판단근거` 를 채운다(과거처럼 `llm_score` 의 reason 직접 사용 금지). 체크포인트 `compa2stage_{담당자}_reason_ckpt.json`. **모든 매칭은 이 경로로 항상 긍정형이 나오도록 되어 있음.** 과거 산출물에 부정 톤이 남아 있으면 `_fix_neg_reasons.py`(부정 톤만 골라 `gen_match_reason` 으로 재작성 → `_build_full_inputs.py` 로 통합best.json 재생성)로 교정. 검증은 `_final_check.py`(부정 톤 매칭근거 0건 확인 포함).
- **상세 매칭 근거(4섹션)**: `gen_explanation` 으로 생성. 섹션 고정 순서·제목:
  `[연관성]` · `[수요기술 사양 적합성]` · `[추천 과제의 우수성]` · `[유사 사례 및 실적]`.
  - 논문/특허 실적이 없으면 `[유사 사례 및 실적]` 에 부재성 문구가 올 수 있음(기존 항목들과 동일한 패턴이므로 스크럽하지 않음 → 문서 전체 일관성 유지).
- **정량 점수(LLM점수)와 근거 텍스트는 분리**된 산출물. 근거 문구 생성이 점수/순위에 영향 주지 않음. 점수는 보고서 표면에 노출하지 않는다.
- 상세 블록 구성: `TopN 제목` → **과제 정보표(7필드, §6-4)** → `적합성 판단: <한 줄>` → `상세 매칭 근거` → 4섹션 문단. (예전의 `수행기관: …` 단독 줄은 정보표의 '과제수행기관' 행으로 흡수·제거. 근거 미생성분이 '과제 개요'로 폴백돼 있던 것도 위 형식으로 교체함.)
- **정렬**: 상세 매칭 근거 4섹션 문단은 **양쪽 맞춤(justify)**.
- **문장 띄어쓰기 보정(필수)**: 마침표 뒤 다음 문장이 공백 없이 붙는 경우(`…습니다.이 과제는`)를 자동 교정 → `compa_match.normalize_spacing()`(한글/`%`/닫음부호 뒤 마침표 다음 공백 삽입, 소수점 `3.2`·버전 `H.264`·`β-1,3` 등은 보존). `gen_explanation`·`gen_match_reason` 반환값에 항상 적용되어 **향후 Qwen 생성분도 자동 보정**. 기존 산출물은 `_fix_spacing.py`(필터 pkl 의 추천근거_상세·판단근거에 적용 → `_build_full_inputs.py` 재생성)로 일괄 교정.
- **문체=반말(한다체)로 통일(필수)**: 상세 매칭 근거는 `~다/~한다/~이다/~된다/~있다` 평서형만 쓰고 `~ㅂ니다/~습니다`(합니다·됩니다·입니다·있습니다·제공합니다·활용합니다·기여합니다·평가됩니다 등) 존댓말 종결은 **하나도** 금지. `build_messages` 의 user 프롬프트에 이 금지어를 예시와 함께 명시(few-shot 존댓말 무시하고 한다체 통일) → 향후 생성분은 반말. 기존 존댓말 산출물은 `_regen_banmal.py`(Qwen 문체 변환: 내용·수치·[섹션]·줄바꿈 유지, 종결어미만 한다체; 잔존 시 `_fix5_banmal.py`·`_fix1_banmal.py`) → `_build_full_inputs.py` 재생성으로 교체.
  - ⚠️ 존댓말 검출 정규식은 `습니다|입니다|됩니다` 만으로는 **`합니다` 계열을 놓친다**. `(?<!아)니다`(‘아니다’만 제외한 모든 `~니다`)를 쓸 것. 검증: `추천근거_상세` 존댓말 잔존 0.
- **메타·회피 문구 금지**: `유사 사례 및 실적` 등에서 `이 섹션은…`, `설명을 마무리한다`, `생략한다`, `제공되지 않아 …작성하지 않는다/서술할 수 없다` 같은 내용 없는 상투·회피 문장 금지. 논문/특허가 없으면 그 사실을 언급하지 말고 **과제의 연구 내용·방법이 수요 해결에 어떻게 활용·연결되는지**를 서술. 교정: `_fix_meta_sections.py`(넓은 메타 패턴)·`_fix_punts.py`/`_fix_punts2.py`(내용 없는 회피형만).

## 5. 데이터 소스 / 필드 매핑

| 필요 값 | 출처 |
|---|---|
| 최종 추천(과제·기관·점수·상세근거) | `COMPA_{담당자}_최종추천.pkl`/`.xlsx`, `COMPA_통합best.json` |
| 후보 pid·수요원문·과제설명문 | `COMPA_후보풀.pkl` |
| 과제 메타(수행기관·논문/특허·상위비율·키워드) | `project_match_data_260612.pkl` (pid 키) |
| 과제설명문·유망성·제출년도·키워드 | `public_RnD_embeddings_pro_260601_with_desc.pkl` |
| 기업설명문(상세근거용) | `company_embeddings_pro_260514_with_desc.pkl` |
| 과제 상세필드(총연구비·기간·분류중·수행기관) | `../apollo/df_project_dataset_260602.pkl` (pid 키) |
| 과학기술표준분류 **대**(1-대) | `../apollo/df_project_all_bizno_260310.pkl` (~86%) |
| 연구개발단계·연구수행주체 | `../apollo/public_RnD_embeddings_pro_260601_with_desc_260708.pkl` |

- **과제수행년도**: `과제설명문` 의 “YYYY년 M월 D일에 시작 … YYYY년 M월 D일에 종료” 에서 **시작·종료 연도** 추출(`제출년도` 단일값은 수행기간과 어긋날 수 있어 미사용). 시작=종료면 단일 연도.

### 5-1. 과제 정보표 7필드 ↔ 소스 컬럼 (pid=과제고유번호 조인)

| 표 항목 | 소스 파일 | 컬럼 | 커버리지 | 비고 |
|---|---|---|---|---|
| 과제고유번호 | 매칭 결과 / dataset | `과제고유번호` | 100% | pid 그 자체 |
| 과제수행기간 | `df_project_dataset_260602.pkl` | `연구시작일`·`연구종료일`(YYYYMMDD) | 100% | `YYYY.MM.DD ~ YYYY.MM.DD` 로 포맷 |
| 과학기술표준분류(중) | `df_project_dataset_260602.pkl` | `과학기술표준분류1-중` | 100% | |
| 과학기술표준분류(대) | `df_project_all_bizno_260310.pkl` | `과학기술표준분류1-대` | ~86% | 없으면 중분류만 표기 |
| 총연구비 | `df_project_dataset_260602.pkl` | `총연구비`(원) | 100% | ≥1억 `n.n억원`, ≥1만 `n만원` 환산 |
| 과제수행기관 | `df_project_dataset_260602.pkl` | `과제수행기관명` | 100% | |
| 연구개발단계 | `public_RnD_embeddings_pro_with_desc_260708.pkl` | `연구개발단계` | 100% | 기초/응용/개발연구/기타 |
| 연구수행주체 | `public_RnD_embeddings_pro_with_desc_260708.pkl` | `연구수행주체` | 100% | 대학/출연연구소/국공립연구소/정부부처(필터됨) |
| 연구책임자 | `public_RnD_PI_260610.pkl` | `연구책임자명` | ~99% | 필터 재매칭본 373/378 |
| 국가연구자번호 | `public_RnD_PI_260610.pkl` | `국가연구자번호` | ~99% | 없으면 행 생략 |

- 위 3개 apollo 파일은 이 저장소 밖(`../apollo/`)에 있으며 대용량(각 3~4GB). 매칭 pid만 필터링해 `scratchpad/pid_fields.json`(pid→필드값 + `(기업,수요기술명,과제명)→pid` 매핑)으로 캐싱한 뒤 docx 패치에 사용.
- `연구단계`·`연구수행주체`는 기존 매칭용 파일(`project_match_data_260612.pkl` 등)엔 없고, **오늘자(260708) 재생성 임베딩 파일에만** 있음.
- ⚠️ **pid 매핑 키는 반드시 `(기업, 수요기술명, 과제명)`을 쓸 것.** `(과제명, 수행기관)`만으로 매핑하면 **동일 과제명이 여러 기업에 서로 다른 과제고유번호로 존재**할 때 충돌해 엉뚱한 pid(→틀린 총연구비·기간·분류·단계·주체)가 들어간다. (실제로 이 오류로 정보표 7건이 잘못 채워졌다가 정정됨.) 과제고유번호의 정답 출처는 매칭 산출물(pkl/`통합best.json`)이며, docx 정보표의 pid를 신뢰하지 말 것.

## 6. 표 서식 규칙 (공통: 고정 레이아웃 `fixed`, 폭 합계 8640 dxa)

### 6-1. 수요 정보 표 (`[수요 N]` 아래, 2열)
- 행: `기업명` / `수요기술 내용` / `수요기술 사양`. **사양이 원본에 비어 있으면 그 행은 생략**(52건 3행, 26건 2행).
- 열 폭: 라벨 **1560** / 값 **7080**. 테두리 단선 `808080`(sz 4, insideH/V 포함).
- 라벨 셀: 음영 `F2F2F2`, 굵게 9pt, 세로 가운데. 값 셀: 원본 문단 복제(줄바꿈·서식 보존).
- **정렬**: `수요기술 내용`·`수요기술 사양` 값 셀은 **양쪽 맞춤(justify)**.

### 6-2. 수요기술 목록 표 (장 시작, 3열)
- 열: `번호` / `수요기술명` / `기업명`.
- 열 폭: **720 / 5040 / 2880** (번호 좁게, 수요기술명 넓게).

### 6-3. 최종 추천 Top5 표 (5열)
- 열: `순위` / `과제명` / `수행기관` / `과제수행년도` / `매칭 근거`.
- 열 폭: **640 / 3360 / 1500 / 1040 / 2100**.
- **정렬**: 제목행 **전체 가운데**. 데이터셀은 `순위`·`수행기관`·`과제수행년도` **가운데**, `매칭 근거` **양쪽 맞춤(justify)**, `과제명` 좌측.
- **과제수행년도 표기**: `시작년도~` (1줄) + `종료년도` (다음 줄) 2줄. 단일 연도는 1줄.

### 6-4. 과제 정보표 (각 TopN 상세 블록, **4열** 라벨|값|라벨|값)
- 위치: TopN 제목 바로 아래. **4열**로 항목 2개를 한 행에 배치해 행 수를 절반으로 축소(항목 홀수면 마지막 우측은 빈칸, 배경 없음).
- 열 폭(docx): 라벨 **2300** / 값 **2020** (×2, 합 8640). 라벨 열은 최장 라벨 `과학기술표준분류(중)`이 **한 줄**에 들어가는 최소 폭. PDF는 라벨 33mm.
- 라벨 셀 음영 `F2F2F2`(PDF `EAF0F7`)·굵게 9pt·네이비. 테두리 `808080`(PDF hairline).
- **항상 8개 항목 고정 순서**(데이터 없으면 `-` 표시, 행 생략 금지): `과제고유번호` · `과제수행기간`(YYYY.MM.DD~ 또는 YYYY년~YYYY년) · `과학기술표준분류(중)` · `연구개발단계` · `과제수행기관` · `연구수행주체` · `연구책임자` · `국가연구자번호`. → 4열 표에서 항상 4행 고정.
- ⚠️ `table_fixed` 는 python-docx 기본 `tblW(w=0 autofit)`·`tblLayout` 을 **먼저 제거**하고 재설정할 것. 중복되면 Word가 첫 autofit 을 따라 고정폭을 무시(라벨 줄바꿈 발생). PDF·docx 모두 이 4열·라벨폭 규칙 동일.

### 6-5. 특허 실적 표 (각 TopN 상세 블록, 상세 매칭 근거 다음)
- **모든 과제에 표시**: 실적 있으면 `▍특허 실적 (N건)` + 8열 표(`구분`·`특허명`·`출원·등록기관`·`국가`·`출원일`·`출원번호`·`등록일`·`등록번호`, 헤더 네이비·흰 글자·얼룩말 행·8pt/docx 7.5pt); **없으면 `▍특허 실적` + "특허 실적 없음"** 한 줄(회색). (실적 있음 117 / 없음 273 = 390)
- **소스**: `../apollo/df_pr_patent_260710_detail.pkl` → `_patent_prep.py` 로 `scratchpad/pid_patents.json` 캐시(매칭 pid 115개, 병합 후 426건).
- **다년도 과제**: 데이터가 `과제고유번호` 로 묶여 있어 **해당 pid 의 전 연도 특허가 자동 포함**.
- **출원/등록 병합**: 같은 특허가 출원 레코드(출원일·출원번호)와 등록 레코드(등록번호·등록일)로 분리되어 있으므로 `(과제, 출원번호)` 로 병합해 한 행으로. **등록 실적 우선**(정렬: 등록 먼저→최신순), 등록 건도 **출원일·출원번호 병기**. 등록 없으면 `구분=출원`.
- 날짜 `YYYYMMDD`→`YYYY.MM.DD`; 일자 `00`이면 `YYYY.MM`, 월도 `00`이면 `YYYY` 로 정리.
- **국가 표기**: 코드가 아니라 **국가명**으로(`_patent_prep.py`의 `country()` 매핑). `KR→한국·US→미국·CN→중국·JP→일본·EP/XU→유럽·XI/WO→PCT`(XI=`PCT/KR…` 출원, XU=EP 출원번호 형식), 미매핑 코드는 원문 유지.
- 특허 표는 길어질 수 있어(최다 21건) 상세 블록 KeepTogether **밖**에 두어 다음 페이지로 흐르도록 허용(1과제 1페이지 규칙의 예외).

## 7. 표기 금지 / 제외

- **'출처' 관련 표기 전면 삭제**: Top 제목의 `(출처:기존/신규/수동)` 접미사, 분야 개요의 `· 출처 …`, `· 후보 출처 태그 —` 범례 줄, `최종 추천 과제 출처 분포 …` 줄 모두 제거.

## 8. 편집 원칙

- 실제 파일 수정 전 **복사본에서 검증**하고, 수정 직전본을 백업(`scratchpad/…BEFORE_*.docx`).
- 서식 신설 시 값을 새로 타이핑하기보다 **정상 항목의 문단/셀 XML을 복제**해 폰트·간격·테두리를 승계.
- 매 편집 후 문단 수 변화·docx zip 무결성(`zipfile.testzip()`)·해당 항목 육안 확인으로 손실 여부 점검.
- **출력 파일명은 매 실행마다 새로 부여**(버전 자동 증가: `..._보고서_v1.docx`, `_v2.docx` …). 이전 산출물이 Word/뷰어에서 열려 있어 파일 잠금(`PermissionError`)이 발생해도 충돌 없이 새 파일을 생성하기 위함. (참고: `gen_report.py` 의 `next_out_path()` — 기존 `_v*.docx` 최대 버전+1.)
- 전체 재생성 스크립트: `gen_report.py`(JSON+pkl+PDF→docx 신규 생성, 출판물 디자인 §9), 검증 `_verify_report.py`(구조·데이터 정합성·서식·폰트·페이지나누기·로고 49항목 자동 점검). 상세근거 1페이지 조판값은 §1-1이 아니라 **§9-5**(세리프 반영) 기준.
- 편집 스크립트(참고): `_patch_docx.py`(근거 채움+헤더), `_strip_source.py`(출처 제거), `_add_disclaimer.py`(경고문), `_add_pagenum.py`(페이지번호), `_demand_table.py`(수요 정보 표), `_add_year_col.py`(수행년도 열), `_fmt_top5.py`(정렬·2줄), `_justify.py`(양쪽 맞춤), `_proj_meta_prep.py`+`_add_proj_table.py`(과제 정보표), `_del_summary.py`(전체 요약 삭제), 근거 생성 `_regen_run.py`·`_regen_shortreason.py`.
- 통합 생성: **`build_report.py`** — 원본(`COMPA_최종Top5_보고서_원본.docx`)에 위 규칙을 순서대로 적용해 최종 보고서를 재현(전체 390건). 기본은 기존 근거 재사용(모델 불필요), `--regen` 시 Qwen3.5-35B-A3B로 근거 재생성.
- 정본 동기화: **`_rebuild_final.py`** — `통합best.json`·담당자 pkl/xlsx의 top5를 보고서 최종 선정으로 교체(식별키 `(기업,수요,과제명)`, pid는 소스에서). **`_fix_projtable_pids.py`** — 과제 정보표 pid 충돌 정정.
- 보조 데이터 캐시: **`_build_full_inputs.py`**(필터 배치→통합best+pid_fields+demand_field), **`_patent_prep.py`**(특허 실적→pid_patents.json), **`rematch_filtered.py`**(필터 재매칭 드라이버).

## 9. 출판물 디자인 (v2~; `gen_report.py`)

전면 신규 생성본에 적용한 출판물 품질 디자인. 데이터·근거 텍스트·페이지 규칙(§1~§8, §1-1)은 유지하고 조판/서식만 상향한다.

### 9-1. 폰트 (에디토리얼 페어링)
- **표제·표·라벨 = `Noto Sans KR`**, **본문 프로즈(개요·상세 매칭 근거·수요 내용/사양/매칭근거) = `Noto Serif KR`**.
- 두 폰트 모두 사용자 Windows 기본 설치(`C:\Windows\Fonts\NotoSans/SerifKR-VF.ttf`) → Word에서 그대로 렌더. 라틴/숫자도 같은 패밀리(ascii·hAnsi·eastAsia·cs 모두 지정)로 통일.
- 검증 렌더(LibreOffice) 시 동일 폰트를 `~/.local/share/fonts`에 넣어 사용자 화면과 일치시킴.

### 9-2. 색 팔레트
- 먹색 `1B2430` · 주색 네이비 `14315C` · 보조 블루 `2C5FA0` · 강조 틸 `0E7C86` · 캡션 `6B7683` · 얇은 괘선 `C9D2DE` · 라벨셀 `EAF0F7` · 짝수행 zebra `F5F8FC`.
- 유의사항 박스: 테두리 `C0392B` / 음영 `FBEEED` / 제목 `A93226` / 본문 `7B241C`(기존 대비 톤 정제).

### 9-3. 구조·요소
- **표지(독립 페이지)**: 영문 키커(레터스페이싱) → 네이비 대제목 → 네이비 굵은 rule → 부제 → **통계 카드 2열**(대상 수요기술 78건 / 추천 과제 390건) → 발행일·생성도구 → 유의사항 박스. (※ '중복 제외 과제' 카드는 표지에서 제외. 중복 제외 수치는 개요 문단에만 유지.)
- **APOLLO 로고(매 페이지)**: `APOLLO 로고.png`의 **위쪽(밝은 배경용) 로고만** 잘라(`make_top_logo()`: 알파 밴드 검출→상단 밴드 crop) 사용. **두 구역(표지·목차 / 본문) 헤더 모두 우측 상단에 삽입**(폭 1.15in) → 표지 포함 모든 페이지 노출. 미디어는 `word/media/image1.png` 1개로 임베드되어 양쪽 헤더가 공유.
- **개요·목차(다음 페이지)**: `▍` 강조바 섹션 라벨. **상세 목차**: 분야(장)별로 그룹핑하고, 각 장 아래 **수요기술별(수요 N · 수요기술명 · 시작 페이지)** 를 모두 명시(장 헤더는 네이비 배경, 우측에 페이지).
  - **PDF(reportlab)**: 페이지 번호를 **2-pass** 로 산출 — pass1 에서 각 장(`_chapter_no`)·수요(`_toc_k`) flowable 의 절대 페이지를 `afterFlowable` 로 수집 → 본문 시작(1장) 기준 상대 페이지(`abs−body_start+1`) 계산, footer 기준 `COVER_PAGES` 도 보정 → pass2 에서 실제 페이지로 목차 렌더. (목차 줄 수는 두 pass 동일하여 페이지 안정.)
  - **docx**: Word **목차 필드**(`TOC \o "1-2" \h \z \u`, 레벨1=장·레벨2=수요) + `settings.updateFields=true`. Word 에서 열면 페이지가 자동 계산·갱신됨(수동 갱신 F9). ※ LibreOffice headless 변환은 목차 필드를 자동 채우지 않음(검증 한계) — 배포용 PDF 는 reportlab 본을 사용하므로 무방.
- **장 오프너**: 네이비 대제목(H1) + 영문 부제(틸) + 네이비 rule + 건수 캡션. 2장부터 새 페이지.
- **러닝 헤더/푸터 (현재 수요기술명 표시)**: 상단 좌측에 `▍` 강조바 + **현재 수요기술명**(그 페이지가 속한 수요), 우측 로고, 하단 얇은 괘선(좌측 끝 틸 강조 틱). 하단 가운데 `— PAGE —`. 구역1(표지·목차)은 헤더/푸터 없음.
  - **docx**: 헤더에 **STYLEREF 필드**(`STYLEREF "Heading 2"`) → 가장 가까운 Heading 2(=수요 제목) 텍스트를 Word/LibreOffice가 페이지마다 자동 표시. 긴 제목은 자동 줄바꿈.
  - **PDF(reportlab)**: 수요기술명 헤더는 **추천 과제 상세(TOP) 페이지에만** 표시. 상세 블록 제목 flowable에 `_demand_hdr=(번호,수요기술명)` 표식 → `afterFlowable` 에서 그 페이지 상단에 `수요 N` 라운드 배지(틸)+수요기술명(네이비) 직접 그림(폭 초과 시 말줄임 …). `onPage` 는 로고·괘선·페이지번호 등 공통 요소만. **새 수요 시작(intro) 페이지에는 수요기술명 헤더를 넣지 않음**(큰 수요 제목과 중복 방지). ※ docx 는 STYLEREF 특성상 intro 페이지에도 현재 수요명이 표시됨(이전 수요명 표시 버그는 없음).
- **수요 제목(H2)**: 틸 `수요 N` 배지 + 네이비 제목 + 하단 얇은 괘선. **과제 제목(H3)**: 네이비 `TOP N` 배지 + 먹색 제목 + 하단 괘선. **적합성 판단**: 블루 배지 태그. **상세근거 4섹션**: 대괄호 제거하고 블루 굵은 라벨 + 세리프 본문.
- **배지 내어쓰기(hanging indent)**: `수요 N`·`TOP N` 배지 뒤 제목이 두 줄 이상이면 둘째 줄이 첫 줄 제목 시작과 **정확히 정렬**되도록 내어쓰기. 들여쓰기 값 = **배지+간격 실측폭을 헤딩마다 계산**(수요 번호 자릿수에 따라 배지 폭이 달라지므로 고정값 금지). PDF는 `pdfmetrics.stringWidth(" 수요 {k} " 등)`, docx는 `PIL.ImageFont(NotoSansKR-Bold).getlength`로 측정(H2 배지 12pt·H3 배지 11pt·간격은 각 base 폰트). 측정 실패 시 근사 폴백.

### 9-4. 표 서식(모던)
- 헤더행 **네이비 solid 음영 + 흰 글자**, 데이터행 **zebra**(짝수행 옅은 배경), 괘선은 얇은 `C9D2DE` 전체 그리드, **셀 여백(tblCellMar)** 부여로 여백 확보. 폭 합계·고정 레이아웃(`fixed`)은 §6 유지.
- Top5·수요목록 표 행에 `cantSplit`(행이 페이지 경계에서 조각나지 않게) → **잔여 슬리버/거의 빈 페이지 방지**.

### 9-5. 1페이지 수렴 재조정(세리프 반영)
- 세리프 본문이 산세리프보다 높이가 커, 과제 상세 조판값을 **본문 8.6pt · 줄간격 1.1 · 섹션간격 2pt · 여백 하 0.7in/좌우 0.9in** 로 재설정(최장 근거 ≈2,300자까지 1페이지 수렴).
- **상세 페이지 세로 리듬(간격 조율)**: 헤더↔본문은 *적당히*(과하지 않게), 블록 내부는 강약을 둔다.
  - 헤더↔제목: docx `top_margin=0.86in`(header_distance 0.35in) / PDF `TM=23mm`·`HDR_Y=PAGE_H−14mm`(≈9mm) — 이전보다 좁힘.
  - 제목(TOP)↔정보표: **좁게**(docx h3 space_after 3pt / PDF 제목 spaceAfter 3 + Spacer 2).
  - 정보표↔적합성판단↔상세매칭근거↔특허실적: **적당히 넓게**(docx space_before 9~11pt / PDF Spacer 9·section before 9~10) — 항목 구분감↑.
  - 4열 정보표로 확보한 세로 여유로 1페이지 수렴 유지.

### 9-6. PDF 생성기 (`gen_report_pdf.py`, reportlab)
- Word/LibreOffice 없이 **동일 데이터·동일 디자인**으로 PDF를 직접 렌더(배포용 PDF의 정식 경로). docx와 팔레트·구조·배지·표·러닝헤더·로고·페이지번호 모두 일치.
- **폰트**: Noto Sans/Serif KR **가변폰트(VF)를 static Regular/Bold로 인스턴싱**해 임베드(`scratchpad/fonts/NotoSans|SerifKR-{Regular,Bold}.ttf`). reportlab 은 OTF(CFF)·가변축 미지원 → fonttools `instantiateVariableFont(wght=400/700)` 로 생성. 표제·표=Sans, 프로즈=Serif.
- **로고**: `apollo_top.png`(상단 밴드 crop)를 표지 우상단 + 본문 러닝헤더 우측에 `canvas.drawImage`(onPage 콜백)로 그림. 러닝헤더 하단 얇은 괘선, 하단 가운데 `— n —`(표지·목차 2p 제외).
- 배포용 PDF는 **이 reportlab 생성본**을 사용(사용자 결정). docx→LibreOffice 변환본은 사용하지 않음.
- 실행: `python gen_report_pdf.py` (환경변수 `COMPA_SCRATCH`·`COMPA_REPORT_JSON`·`COMPA_PDF_OUT` 로 경로 지정 가능).

### 9-7. docx 렌더 검증 (LibreOffice)
- 이 저장소 밖 폰트/렌더 확인용으로 LibreOffice(headless) 사용 가능: `soffice --headless --convert-to pdf --outdir <dir> <docx>` → 페이지 추출 후 `sips`로 PNG 렌더해 육안 검증. (docx 자체 배포는 하지 않고 검증에만 사용.)
- 수요 요약(정보표+Top5)도 밀도 상향(정보표 9pt/줄간격 1.28, Top5 8.8pt/줄간격 1.18)해 **near-empty 페이지 0** 달성.
- 검증(`_verify_report.py`, PASS 45): 과제 상세 overflow 0/390, near-empty 페이지 0, 데이터 정합성·서식·폰트·페이지나누기 전 항목 통과.

## 10. 변경 이력 (2026-07 필터 재매칭 + 서식·문체 정제 세션)

최근 세션에서 적용한 주요 변경(최신 산출물 `COMPA_필터전체_보고서.pdf`/`.docx`, 정본 `COMPA_통합best.json`):

1. **필터 재매칭(§0)**: 제출년도 ≥2020 ∧ 연구수행주체 {대학·출연연·국공립연·정부부처} 조건으로 78수요 전체 35B 재매칭. 위배 0건, 390건 재선정. (`rematch_filtered.py` → `_build_full_inputs.py`)
2. **연구책임자 정보 추가(§5-1, §6-4)**: 과제 정보표에 `연구책임자`·`국가연구자번호` 행 추가(소스 `../apollo/public_RnD_PI_260610.pkl`).
3. **과제 정보표 4열화(§6-4)**: 세로 2열(≤8행) → 4열(라벨|값|라벨|값)로 행 수 절반 축소. 라벨열 폭은 `과학기술표준분류(중)` 한 줄 기준.
4. **특허 실적 표(§6-5)**: 각 과제에 `▍특허 실적` 표(등록 우선·출원 병기, 다년도 전 연도). 없으면 "특허 실적 없음" 표기. (`_patent_prep.py`)
5. **PDF 출판물 재현(§9-6)**: reportlab로 docx와 동일 디자인 PDF 직접 생성(Noto Sans/Serif KR static, 로고·배지·표). 배포용 PDF는 이 생성본.
6. **러닝헤더=현재 수요기술명(§9-3)**: 상세 페이지 상단에 `수요 N`+수요기술명. PDF는 상세 페이지에만(새 수요 intro 페이지는 미표시), docx는 STYLEREF.
7. **세로 리듬·간격(§9-5)**: 헤더↔제목·제목↔정보표는 좁게, 정보표↔적합성↔상세근거↔특허는 적당히 넓게.
8. **배지 내어쓰기(§9-3)**: `수요 N`·`TOP N` 뒤 제목의 둘째 줄이 첫 줄 제목 시작과 정확히 정렬(배지+간격 실측폭 헤딩별 계산).
9. **긍정형 매칭 근거(§4)**: 표 '매칭 근거'(판단근거)를 `compa_match.gen_match_reason`(긍정 few-shot)으로 생성하도록 `match_for` 수정 → 향후 항상 긍정형. 기존 부정 톤 13건은 `_fix_neg_reasons.py`로 교정. 부정 톤 0.
10. **상세 목차(§9-3)**: 분야(장)별로 그룹핑하고 각 수요기술별 제목+시작 페이지 명시. PDF는 **2-pass**로 페이지 산출(목차↔본문 페이지 전수 일치 검증), docx는 Word **TOC 필드**(레벨1-2). 장별 건수 표기는 제외.
11. **과제 정보표 8필드 고정(§6-4)**: 항상 `과제고유번호·과제수행기간·과학기술표준분류(중)·연구개발단계·과제수행기관·연구수행주체·연구책임자·국가연구자번호` 순서, 데이터 없으면 `-`(행 생략 금지).
12. **특허 국가명(§6-5)**: 코드→국가명(`KR→한국·US→미국·CN→중국·JP→일본·EP/XU→유럽·XI/WO→PCT`).
13. **제목 볼드**: 표지 타이틀·수요기술명(H2)·TOP 과제명(H3) 볼드(Sans-B/`bold=True`).
14. **문장 띄어쓰기 보정(§4)**: 마침표 뒤 공백 누락 자동 교정 `compa_match.normalize_spacing`(소수점·버전 보존), `gen_explanation`/`gen_match_reason` 항상 적용. 기존분 `_fix_spacing.py`.
15. **문체=반말(한다체) 통일(§4)**: 상세 매칭 근거를 반말로. 프롬프트에 존댓말 금지어(`합니다`류 포함) 명시. 기존분 `_regen_banmal.py`(+`_fix5/_fix1`) 교체. 존댓말 검출은 `(?<!아)니다`. 존댓말 0.
16. **메타·회피 문구 제거(§4)**: `유사 사례 및 실적` 등의 `이 섹션은…/생략한다/제공되지 않아…서술할 수 없다` 상투·회피 문장을 기술 연결성 서술로 재생성(`_fix_meta_sections.py`·`_fix_punts.py`·`_fix_punts2.py`·`_fix_usecase.py`). 회피/메타 0.

> 검증: **`_final_check.py`**(27항목: 구조·필터·특허·볼드·4열표·부정톤0·존댓말0·pid정합성·무결성) — PASS.
> 스크립트: `rematch_filtered.py`, `_build_full_inputs.py`, `_patent_prep.py`, `gen_report.py`(docx), `gen_report_pdf.py`(PDF), `_rebuild_final.py`, `_fix_projtable_pids.py`, 교정용 `_fix_*` / `_regen_banmal.py`. 데이터·산출물(pkl/xlsx/docx/pdf/폰트)은 `.gitignore` 제외.
