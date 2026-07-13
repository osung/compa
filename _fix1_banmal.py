# -*- coding: utf-8 -*-
"""존댓말이 계속 남는 1행을 섹션별로 나눠 반말 변환(확실)."""
import re, pandas as pd, compa_match as cm
PKL = "COMPA_필터21_78_최종추천.pkl"
JON = re.compile(r"습니다|입니다|됩니다|있습니다|합니다|봅니다|집니다|립니다|칩니다|십니다")
SYS = ("다음 한 단락의 의미·수치·고유명사를 그대로 두고, 모든 문장 종결을 예외 없이 '~다/~한다/~이다/~된다/~있다' "
       "평서형(한다체)으로 바꾼다. 존댓말('~습니다/~입니다/~됩니다/~있습니다/~합니다')이 하나도 남으면 안 된다. "
       "머리말 없이 변환된 단락만 출력한다.")

def conv(t):
    best = t
    for temp in (0.0, 0.4, 0.7):
        o = cm.normalize_spacing(cm.stream_explanation(
            [{"role": "system", "content": SYS}, {"role": "user", "content": t}],
            max_tokens=900, temperature=temp, top_p=1.0).strip())
        o = re.sub(r"^(변환\s*결과|출력)\s*[:：]\s*", "", o)
        if o and len(o) > len(t) * 0.6:
            if not JON.search(o): return o
            if len(JON.findall(o)) < len(JON.findall(best)): best = o
    return best

def main():
    df = pd.read_pickle(PKL)
    idx = next(i for i, r in df.iterrows()
               if str(r["과제고유번호"]) == "1711120839" and str(r["번호"]) == "76")
    src = str(df.at[idx, "추천근거_상세"])
    parts = re.split(r"(\[[^\]]+\])", src)   # ['', '[연관성]', ' body', '[..]', ' body', ...]
    cm.load_model_blocking(progress_cb=lambda m: print(" ", m, flush=True))
    out = parts[0]
    for j in range(1, len(parts), 2):
        tag = parts[j]; body = parts[j + 1] if j + 1 < len(parts) else ""
        nb = conv(body) if JON.search(body) else body
        out += tag + nb
        print(f"  {tag} 존댓말 {len(JON.findall(nb))}개", flush=True)
    out = cm.normalize_spacing(out)
    df.at[idx, "추천근거_상세"] = out
    df.to_pickle(PKL); df.to_excel(PKL.replace(".pkl", ".xlsx"), index=False)
    print("잔존:", len(JON.findall(out)), flush=True)

if __name__ == "__main__":
    main()
