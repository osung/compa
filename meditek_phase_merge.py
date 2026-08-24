# -*- coding: utf-8 -*-
"""다년차 과제 통합 — 같은 과제의 연차 보고를 한 건으로 묶는다.

NTIS 는 다년도 과제의 **연차 보고마다 별도의 과제고유번호**를 부여한다. 그래서 같은
과제가 연차 수만큼 행으로 등재되고, 연차마다 누적 성과가 달라 특허·논문 건수와
유망성점수가 갈린다. 보고서에서 같은 과제명이 다른 특허 실적을 달고 두 번 나오면
오류로 읽히므로 하나로 통합한다.

  예) '영상진단 의료기기 탑재용 AI 진단 기술 개발' (경북대학교)
      1415169720(2020) · 1415173238(2021) · 1415179472(2022) · 1415187778(2023)

그룹 키 = 과제명 정규화 + **공급기관 계열**
  · 과제명만으로 묶으면 안 된다 — '실험실 특화형 창업선도대학 사업'(울산과학기술원 /
    충북대학교), '핵심연구지원센터조성지원과제'(울산대·경북대·인제대, 전부 2022년)처럼
    여러 기관이 같은 이름의 국가사업을 각각 수행한다. 이들은 별개 과제다.
  · 과제수행기관명을 그대로 넣어도 안 된다 — 같은 과제인데 연차마다 기관 표기가 바뀐다
    (서울아산병원↔아산사회복지재단, 고려대학교 안암병원↔고려대학의과대학부속병원,
     가톨릭대학교↔가톨릭대학교 성의캠퍼스). 실측 64개 그룹이 이 경우였다.
  · 과제고유번호 접두로도 못 묶는다 — 2024년에 체계가 개편돼 같은 과제가
    1465030925(2020) → 2460000169(2024) 처럼 바뀐다.
  → 이미 정의해 둔 공급기관 계열(meditek_supply_orgs.include_map)이 정확히 필요한
    정규화를 해 준다.

대표 행 = **제출년도가 가장 늦은 연차**(성과가 가장 많이 누적된 보고). 과제명·설명문·
유망성·기간 등 속성은 대표 행에서 가져오고, **특허는 전 연차의 합집합**을 쓴다.
"""
import json
import os

import compa_match as cm


def group_key(name, supplier):
    return f"{cm.title_key(name)}|{str(supplier).strip()}"


def build_groups(pids, names, suppliers, years):
    """→ (rep_of, members)

    rep_of : {과제고유번호: 대표 과제고유번호}
    members: {대표 과제고유번호: [연차 과제고유번호…]}  (제출년도 오름차순)
    """
    buck = {}
    for p, n, s, y in zip(pids, names, suppliers, years):
        try:
            yy = int(y)
        except (TypeError, ValueError):
            yy = 0
        buck.setdefault(group_key(n, s), []).append((yy, str(p)))
    rep_of, members = {}, {}
    for _, lst in buck.items():
        lst.sort()                               # 연도 오름차순, 동년은 번호순
        rep = lst[-1][1]                         # 가장 늦은 연차를 대표로
        members[rep] = [p for _, p in lst]
        for _, p in lst:
            rep_of[p] = rep
    return rep_of, members


def save(members, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(members, f, ensure_ascii=False)
    return path


def load(path):
    if not path or not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def stats(members):
    n_multi = sum(1 for v in members.values() if len(v) > 1)
    n_row = sum(len(v) for v in members.values())
    mx = max((len(v) for v in members.values()), default=0)
    return dict(groups=len(members), rows=n_row, multi=n_multi,
                collapsed=n_row - len(members), max_size=mx)
