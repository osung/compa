import ast
from pathlib import Path

import numpy as np
import pandas as pd
from numpy.linalg import norm
#from sentence_transformers import SentenceTransformer
import torch
from transformers import AutoTokenizer, AutoModel
import time


def load_data(data_path="./data/data.csv"):
    return pd.read_csv(data_path)


def normalize_company_name(name):
    name = str(name).strip()

    name = name.replace(" ", "")
    name = name.replace("(주)", "")
    name = name.replace("㈜", "")
    name = name.replace("주식회사", "")

    return name

def select_company_matches(df, company_name, top_n):
    target_name = normalize_company_name(company_name)

    df = df.copy()
    df["_company_name_norm"] = df["company_name"].apply(normalize_company_name)

    matched = (
        df[df["_company_name_norm"] == target_name]
        .drop(columns=["_company_name_norm"])
        .head(top_n)
        .copy()
    )

    if matched.empty:
        raise ValueError(f"'{company_name}'에 해당하는 데이터가 없습니다.")

    return matched


def ensure_list(x):
    if isinstance(x, list):
        return x

    if isinstance(x, np.ndarray):
        return x.tolist()

    if x is None or (isinstance(x, float) and pd.isna(x)):
        return []

    if isinstance(x, str):
        x = x.strip()
        if not x:
            return []
        try:
            return ast.literal_eval(x)
        except Exception:
            return [x]

    return [x]


def cosine_similarity(a, b):
    a = np.array(a, dtype=float)
    b = np.array(b, dtype=float)

    denom = norm(a) * norm(b)
    if denom == 0:
        return 0.0

    return float(np.dot(a, b) / denom)

def mean_pooling(last_hidden_state, attention_mask):
    mask = attention_mask.unsqueeze(-1).expand(last_hidden_state.size()).float()
    return (last_hidden_state * mask).sum(1) / mask.sum(1)


def embed_texts(texts, model, tokenizer, device):
    encoded = tokenizer(
        texts,
        padding=True,
        truncation=True,
        return_tensors="pt",
    ).to(device)

    with torch.no_grad():
        outputs = model(**encoded)

    embeddings = mean_pooling(
        outputs.last_hidden_state,
        encoded["attention_mask"],
    )

    embeddings = torch.nn.functional.normalize(
        embeddings,
        p=2,
        dim=1,
    )

    return embeddings.cpu().numpy()


def filter_conduct_list(row_embed, conduct_list, threshold=0.5, top_n=3, dedup_key=None):
    conduct = ensure_list(conduct_list)
    if not conduct:
        return None

    embed = ensure_list(row_embed)
    if not embed:
        return None

    result = []

    for item in conduct:
        if not isinstance(item, dict):
            continue

        item_embed = item.get("embedding")
        if item_embed is None:
            continue

        item_embed = ensure_list(item_embed)
        if not item_embed:
            continue

        try:
            sim = cosine_similarity(embed, item_embed)
        except Exception:
            continue

        if sim >= threshold:
            result.append((sim, item))

    if not result:
        return None

    result = sorted(result, key=lambda x: x[0], reverse=True)

    if dedup_key is not None:
        seen = set()
        deduped = []

        for sim, item in result:
            val = item.get(dedup_key)

            if val in seen:
                continue

            seen.add(val)
            deduped.append((sim, item))

        result = deduped

    result = [item for _, item in result[:top_n]]
    return result if result else None


def remove_embedding(x):
    x = ensure_list(x)
    cleaned = []

    for item in x:
        if isinstance(item, dict):
            item = item.copy()
            item.pop("embedding", None)

        cleaned.append(item)

    return cleaned if cleaned else None


def parse_list(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or x == "":
        return []

    if isinstance(x, list):
        return x

    if isinstance(x, str):
        x = x.strip()
        if x == "":
            return []
        try:
            return ast.literal_eval(x)
        except Exception:
            return []

    return []


def remove_duplicate_patents(x):
    if not isinstance(x, list):
        return x

    seen = set()
    unique = []

    for p in x:
        key = str(p)
        if key not in seen:
            seen.add(key)
            unique.append(p)

    return unique


_DEFAULT_NOUN_FN = "UNSET"


def _get_default_noun_fn():
    """기본 명사 추출기 (lazy, 1회 구성). 분석기 미설치 시 None → 원문 폴백.

    과제/기업 임베딩이 '키워드(명사) ;결합' 문자열로 생성되므로, 특허도 명사를 추출해
    ';' 로 이어 임베딩하는 것이 표현공간 정합성이 높다(eval_patent_embed_modes.py 검토).
    프로젝트 키워드 파이프라인과 동일 계열인 MeCab 우선, 없으면 kiwipiepy 사용.
    """
    global _DEFAULT_NOUN_FN
    if _DEFAULT_NOUN_FN != "UNSET":
        return _DEFAULT_NOUN_FN

    fn = None
    try:
        from konlpy.tag import Mecab
        _m = Mecab()
        fn = lambda title: [w for w in _m.nouns(str(title)) if len(w) >= 2]
    except Exception:
        try:
            from kiwipiepy import Kiwi
            _kiwi = Kiwi()
            fn = lambda title: [t.form for t in _kiwi.tokenize(str(title))
                                if t.tag.startswith("NN") and len(t.form) >= 2]
        except Exception:
            fn = None
    _DEFAULT_NOUN_FN = fn
    return _DEFAULT_NOUN_FN


def _patent_embed_text(title, embed_mode, noun_fn):
    """특허 임베딩에 넣을 텍스트 구성.

    embed_mode='raw'  → 특허명 원문 그대로
    embed_mode='noun' → noun_fn 으로 주요 명사를 추출해 ';' 로 이어 붙임
                        (추출 실패/빈 결과면 원문으로 폴백)
    """
    if embed_mode == "noun" and noun_fn is not None:
        try:
            nouns = noun_fn(title)
        except Exception:
            nouns = None
        if nouns:
            return ";".join(nouns)
    return title


def build_company_patent_sim(
    company_patent,
    project_embed,
    model,
    tokenizer=None,
    device=None,
    threshold=0.5,
    top_n=3,
    min_for_topk=10,
    embed_mode="raw",
    noun_fn=None,
    show_progress=False,
    progress_desc=None,
):
    """추천 과제와 유사한 회사 보유특허를 선별.

    - 보유특허가 min_for_topk(기본 10) 미만이면 유사도 선별 없이 '전체'를 반환한다.
    - min_for_topk 이상이면 특허명을 과제 임베딩(project_embed, pro-sroberta 공간)과
      같은 모델(model: pro-sroberta SentenceTransformer)로 임베딩해 cosine 유사도
      상위 top_n(기본 3) 개만 반환한다(threshold 미만 제외).
    - embed_mode: 'raw'(특허명 원문) | 'noun'(주요 명사 ';' 결합) — noun_fn 필요.

    ※ project_embed 와 특허 임베딩은 반드시 동일 모델(pro-sroberta)이어야 한다.
      과거 bge-m3(1024) vs pro-sroberta(768) 혼용으로 차원 불일치가 있었음.
    """
    patents = parse_list(company_patent)
    patents = remove_duplicate_patents(patents)

    patent_texts = []
    for p in patents:
        text = str(p).strip()
        if not text or text.lower() in {"nan", "none", "null"}:
            continue
        patent_texts.append(text)

    if not patent_texts:
        return None

    # 보유특허가 적으면(<min_for_topk) 선별 없이 전체를 LLM 에 제공
    if len(patent_texts) < min_for_topk:
        return patent_texts

    project_embed = ensure_list(project_embed)
    if not project_embed:
        return None

    embed_inputs = [
        _patent_embed_text(t, embed_mode, noun_fn) for t in patent_texts
    ]
    if show_progress:
        # 콘솔에 tqdm progress bar 표시(라벨: progress_desc). 배치 단위로 임베딩.
        from tqdm.auto import tqdm
        _BATCH = 64
        _chunks = []
        for _i in tqdm(range(0, len(embed_inputs), _BATCH),
                       desc=(progress_desc or "유사도 top-3 임베딩"),
                       unit="batch", leave=False):
            _chunks.append(model.encode(embed_inputs[_i:_i + _BATCH],
                                        show_progress_bar=False))
        patent_embeds = np.vstack(_chunks) if _chunks else np.empty((0,))
    else:
        patent_embeds = model.encode(embed_inputs, show_progress_bar=False)

    result = []
    for patent_text, patent_embed in zip(patent_texts, patent_embeds):
        try:
            sim = cosine_similarity(project_embed, patent_embed)
        except Exception:
            continue
        if sim >= threshold:
            result.append((sim, patent_text))

    if not result:
        return None

    result = sorted(result, key=lambda x: x[0], reverse=True)
    output = [patent_text for sim, patent_text in result[:top_n]]
    return output if output else None


def load_embedding_model(model_name="pro-sroberta"):
    """특허 임베딩 모델 로드.

    과제/기업 임베딩(project_norm_embed)이 pro-sroberta(768차원)로 생성되므로,
    동일 공간 비교를 위해 특허도 pro-sroberta SentenceTransformer 로 임베딩한다.
    반환 (model, tokenizer=None, device) — 기존 호출부 시그니처 호환용.
    """
    from sentence_transformers import SentenceTransformer

    device = "cuda" if torch.cuda.is_available() else "cpu"
    model = SentenceTransformer(model_name, device=device)
    return model, None, device

def add_company_patent_sim(
    tmp,
    model,
    tokenizer,
    device,
    threshold=0.5,
    top_n=3,
    min_for_topk=10,
    embed_mode="noun",
    noun_fn=None,
):
    if "company_patent" not in tmp.columns:
        return tmp

    if "project_norm_embed" not in tmp.columns:
        return tmp

    # noun 모드인데 추출기가 안 주어지면 기본 추출기(MeCab→kiwi, 없으면 원문 폴백) 사용
    if embed_mode == "noun" and noun_fn is None:
        noun_fn = _get_default_noun_fn()

    tmp["company_patent_sim"] = tmp.apply(
        lambda row: build_company_patent_sim(
            company_patent=row["company_patent"],
            project_embed=row["project_norm_embed"],
            model=model,
            tokenizer=tokenizer,
            device=device,
            threshold=threshold,
            top_n=top_n,
            min_for_topk=min_for_topk,
            embed_mode=embed_mode,
            noun_fn=noun_fn,
        ),
        axis=1,
    )

    return tmp

def preprocess(
    tmp,
    embed_model=None,
    embed_tokenizer=None,
    embed_device=None,
):
    tmp = tmp.copy()

    if "project_name" not in tmp.columns and "과제명" in tmp.columns:
        tmp["project_name"] = tmp["과제명"]

    if "conduct_list_project" in tmp.columns and "company_norm_embed" in tmp.columns:
        tmp["conduct_list_project"] = tmp.apply(
            lambda row: filter_conduct_list(
                row["company_norm_embed"],
                row["conduct_list_project"],
                threshold=0.5,
                top_n=3,
                dedup_key="project_id",
            ),
            axis=1,
        )

    if "conduct_list_company" in tmp.columns and "project_norm_embed" in tmp.columns:
        tmp["conduct_list_company"] = tmp.apply(
            lambda row: filter_conduct_list(
                row["project_norm_embed"],
                row["conduct_list_company"],
                threshold=0.5,
                top_n=3,
                dedup_key="company_id",
            ),
            axis=1,
        )

    for col in ["conduct_list_company", "conduct_list_project"]:
        if col in tmp.columns:
            tmp[col] = tmp[col].apply(remove_embedding)

    if "company_patent" in tmp.columns:
        tmp["company_patent"] = (
            tmp["company_patent"]
            .apply(parse_list)
            .apply(remove_duplicate_patents)
        )

    if embed_model is not None and embed_tokenizer is not None and embed_device is not None:
        tmp = add_company_patent_sim(
            tmp,
            model=embed_model,
            tokenizer=embed_tokenizer,
            device=embed_device,
        )

    if "patent" in tmp.columns:
        tmp["patent"] = (
            tmp["patent"]
            .apply(parse_list)
            .apply(remove_duplicate_patents)
        )

    tmp = tmp.drop(
        columns=["company_norm_embed", "project_norm_embed"],
        errors="ignore",
    )

    return tmp

'''
def prepare_data(
    company_name,
    top_n,
    data_path="./data/data.csv",
    save_path="./data/tmp.csv",
    embed_model=None,
    embed_tokenizer=None,
    embed_device=None,
):
    print('데이터 전처리 시작')
    t_total_start = time.perf_counter()

    df = load_data(data_path)

    tmp = select_company_matches(
        df=df,
        company_name=company_name,
        top_n=top_n,
    )

    tmp = preprocess(
        tmp,
        embed_model=embed_model,
        embed_tokenizer=embed_tokenizer,
        embed_device=embed_device,
    )

    save_path = Path(save_path)
    save_path.parent.mkdir(parents=True, exist_ok=True)

    tmp.to_csv(
        save_path,
        index=False,
        encoding="utf-8-sig",
    )

    t_total_end = time.perf_counter()
    elapsed = t_total_end - t_total_start
    print(f"[완료] 데이터 전처리 전체: {format_time(elapsed)}")

    return tmp
'''