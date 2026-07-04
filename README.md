# compa

COMPA 진성수요(기업 기술 수요) ↔ 국가 R&D 과제 **매칭 파이프라인**.

각 기업 수요 1건에 대해 키워드 추출 → 임베딩 유사도(SBERT) → LLM 적합도 평가 →
LLM 상세 추천 근거 생성을 거쳐, 가장 적합한 국가 R&D 과제 Top-N 을 추천한다.

## 파이프라인

기업 수요 1건(= 진성수요 시트 한 행)에 대해:

1. **키워드 추출** — 수요기술명·내용·사양 → LLM 으로 핵심 기술 키워드 10~20개
2. **임베딩** — `pro-sroberta` 로 키워드 문자열 인코딩 → L2 정규화 (query 벡터)
3. **단일 SBERT 매칭** — query 와 과제 임베딩의 코사인 유사도 → 상위 `TOPK`(기본 100)
4. **LLM 적합도** — 후보 각각을 0~100 으로 정량 평가(배치) → 상위 `FINAL`(기본 5)
5. **상세 근거** — 최종 `FINAL` 개에 4섹션 상세 추천 근거 생성
6. **산출/후처리** — 최종추천 `xlsx`/`pkl` + 15컬럼 정리본 + 기업별 탭본

- 임베딩: `pro-sroberta` (SentenceTransformers, CPU)
- LLM: Qwen 계열 — **MLX**(Apple Silicon) / **vLLM**(CUDA) 자동 선택

## 구성

| 파일 | 설명 |
|------|------|
| `compa_match.py` | **단일 파일 실행 스크립트.** 아래 모든 모듈을 하나로 합친 self-contained 버전 |
| `compa_colab.ipynb` | Google Colab(A100 + vLLM) 실행 노트북. Qwen3.6-27B / 35B-A3B 선택 |
| `match_compa_2stage.py` | (원본) 진입점 — 단일 SBERT + LLM 적합도 파이프라인 |
| `match_compa_list.py` | (원본) LLM 적합도 점수 / 상세근거 생성 함수 |
| `matching_viewer_explain.py` | (원본) LLM 백엔드(MLX/vLLM) + 프롬프트 |
| `user_input_keywords.py` | (원본) 규칙기반 명사 키워드 추출(LLM 폴백) |
| `postprocess_compa.py`, `split_by_company.py` | (원본) 정리본 / 기업별 탭본 후처리 |
| `company_to_project/`, `project_to_company/` | (원본) 방향별 프롬프트·데이터 파이프라인 |

`compa_match.py` 는 원본 8개 모듈의 프롬프트·로직을 그대로(byte-exact) 인라인한 것으로,
동작이 동일하다. 새로 시작한다면 `compa_match.py` 하나만 있으면 된다.

## 데이터 파일 (Git 제외 — 별도 배포)

대용량 데이터·모델·산출물은 `.gitignore` 로 제외된다. 실행하려면 스크립트와 같은
디렉토리(또는 Colab 은 지정한 Drive 폴더)에 다음이 있어야 한다:

```
public_RnD_embeddings_pro_260601_with_desc.pkl   # 국가 R&D 과제 임베딩
company_embeddings_pro_260514_with_desc.pkl      # 기업 설명문(상세근거용)
project_match_data_260612.pkl                     # 과제 수행기관/논문·특허 성과
pro-sroberta/                                     # SentenceTransformer 모델 폴더
COMPA_진성수요_원본.xlsx                            # 입력(수정하지 않음)
```

## 실행 (로컬)

```bash
pip install numpy pandas openpyxl sentence-transformers
# LLM 백엔드: Apple Silicon → pip install mlx-lm / CUDA → pip install vllm

# 기본 담당자(이중연)
python compa_match.py

# 특정 담당자 / 전체
python compa_match.py --assignee 이중연,김민주
python compa_match.py --all

# 빠른 테스트(수요 1건), 상세근거 생략, 키워드만
python compa_match.py --assignee 이중연 --limit 1
python compa_match.py --all --no-explain
python compa_match.py --all --keywords-only
```

주요 옵션: `--topk`(SBERT 후보 수, 기본 100), `--final`(최종 추천 수, 기본 5),
`--exclude`(제외 담당자), `--limit`(담당자별 수요 수 제한).

입출력 디렉토리는 환경변수로 바꿀 수 있다:
`COMPA_DATA_DIR`(입력 위치), `COMPA_OUT_DIR`(산출물·체크포인트 위치).

## 실행 (Google Colab · A100 · vLLM)

`compa_colab.ipynb` 를 Colab 에서 열고:

1. 런타임 → 하드웨어 가속기 **A100 GPU** 선택
2. **① 설정** 셀에서 Drive 경로·`MODEL_CHOICE`·실행 대상 지정
3. 위에서부터 실행 (또는 런타임 → 모두 실행)

선택 가능한 모델 프리셋:

| MODEL_CHOICE | HF repo | 대략 크기 | 권장 |
|---|---|---|---|
| `27B` | `Qwen/Qwen3.6-27B-FP8` | ~27GB | 40GB A100 |
| `35B-A3B` | `Qwen/Qwen3.6-35B-A3B-FP8` | ~35GB | 80GB A100 |
| `35B-A3B-int4` | `palmfuture/Qwen3.6-35B-A3B-GPTQ-Int4` | ~18GB | 40GB A100 |

## 체크포인트 / 재개

키워드·LLM점수·상세근거를 담당자별 체크포인트(JSON)로 `COMPA_OUT_DIR` 에 저장한다.
중간에 중단(세션 끊김/OOM)되어도 **다시 실행하면 이미 처리한 수요는 LLM 재호출 없이
건너뛴다.** Colab 에서는 `COMPA_OUT_DIR` 을 Google Drive 로 지정하면 세션이 끊겨도
이어서 실행할 수 있다.

## 산출물 (담당자 `{a}` 별)

```
COMPA_{a}_키워드.xlsx                  # 키워드 추출 결과
COMPA_{a}_최종추천.xlsx / .pkl         # 최종 Top-N 추천(점수·근거·상세근거)
COMPA_{a}_최종추천_보완.xlsx           # 15컬럼 정리본
COMPA_{a}_최종추천_보완_기업별.xlsx    # 기업별 탭 분리본
```
