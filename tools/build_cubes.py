#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_cubes.py — /tmp/full.json (build_data.py 산출물) → 대시보드용 다중 큐브

용량 문제 해결을 위해 3개 큐브로 분리합니다.
  core          : (y,m,sd,ch,grp,item,corp)   — 상품축 전부 + 거래선 + 추세탭. 전 연도.
  year_YYYY_Hn  : 위 + store                  — 조직축 조회용. 연도ㆍ반기별 분리.

index.html 은 데이터를 내장하지 않고 두 파일을 fetch 로 읽습니다.

금액은 천원 단위 정수로 저장(÷1000 반올림)하여 용량을 추가로 줄입니다.
"""
import json, os, sys, gzip
from collections import defaultdict

SRC = sys.argv[1] if len(sys.argv) > 1 else "/tmp/full.json"
OUT = sys.argv[2] if len(sys.argv) > 2 else "./gh/data"

Y, M, SD, CH, GRP, MD, IT, CP, ST, AMT, QTY, EFF, SUB = range(13)
K = 1000  # 원 → 천원


def agg(rows, keys):
    """keys 기준으로 amt/qty/eff/sub 합산 (금액은 천원 단위)"""
    d = defaultdict(lambda: [0, 0, 0, 0])
    for r in rows:
        k = tuple(r[i] for i in keys)
        v = d[k]
        v[0] += r[AMT]; v[1] += r[QTY]; v[2] += r[EFF]; v[3] += r[SUB]
    out = []
    for k, v in d.items():
        out.append(list(k) + [round(v[0] / K), v[1], round(v[2] / K), round(v[3] / K)])
    out.sort()
    return out


def write(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)  # NaN 은 JSON 표준이 아님
    raw = os.path.getsize(path)
    with open(path, "rb") as f:
        gz = len(gzip.compress(f.read(), 6))
    print(f"  {os.path.basename(path):<22} {raw/1024/1024:>7.1f}MB  (gzip {gz/1024/1024:>5.1f}MB)")
    return raw, gz


def main():
    print(f"로드: {SRC}")
    d = json.load(open(SRC, encoding="utf-8"))
    rows = d["rows"]
    print(f"  원본 {len(rows):,}행")
    os.makedirs(OUT, exist_ok=True)

    years = sorted({r[Y] for r in rows})

    # ── core: 전 연도, store 없음 ──
    core = agg(rows, [Y, M, SD, CH, GRP, IT, CP])
    print(f"\ncore {len(core):,}행")
    core_obj = {
        "meta": d["meta"],
        "dims": d["dims"],
        "masters": d["masters"],
        "cols": ["y", "m", "sd", "ch", "grp", "item", "corp", "amt", "qty", "effAmt", "sub"],
        "unit": "천원",
        "rows": core,
    }
    write(os.path.join(OUT, "core.json"), core_obj)

    # ── 연도ㆍ반기별 전체그레인 (조직축 조회용) ──
    # GitHub 웹 업로드 제한(25MB)과 로딩 체감을 고려해 반기 단위로 쪼갠다.
    # 큐브를 차원 기준으로 더 쪼개면 "품목필터 + 거래선축 + 지부필터" 같은 조합을
    # 정확히 계산할 수 없으므로, 그레인은 유지하고 기간으로만 분할한다.
    print(f"\n연도ㆍ반기별 전체그레인 (조직축 조회용)")
    tot_raw = tot_gz = 0
    index = {"years": years, "files": {}}
    for y in years:
        yr = [r for r in rows if r[Y] == y]
        for h_, ms in (("H1", range(1, 7)), ("H2", range(7, 13))):
            sub = [r for r in yr if r[M] in ms]
            if not sub:
                continue
            a = agg(sub, [M, SD, CH, GRP, IT, CP, ST])
            obj = {
                "y": y, "half": h_,
                "cols": ["m", "sd", "ch", "grp", "item", "corp", "store",
                         "amt", "qty", "effAmt", "sub"],
                "rows": a,
            }
            fn = f"year_{y}_{h_}.json"
            r_, g_ = write(os.path.join(OUT, fn), obj)
            tot_raw += r_; tot_gz += g_
            index["files"].setdefault(str(y), []).append(fn)
    write(os.path.join(OUT, "year_index.json"), index)

    print(f"\n조직축 합계 {tot_raw/1024/1024:.1f}MB (gzip {tot_gz/1024/1024:.1f}MB)")
    print("→ 기본 화면(상품축ㆍ추세탭)은 core.json만으로 동작")
    print("→ 조직축 선택 시 해당 연도 + 전년 + 전전년의 필요 반기 파일만 추가 로드")


if __name__ == "__main__":
    main()
