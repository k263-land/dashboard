#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — 일보DB(CSV) + 기초DB(XLSX) → 대시보드용 JSON 생성

인수인계서 §2 "채널·그룹 판정 규칙"을 그대로 구현한 스크립트입니다.
집계 단계에서 채널ㆍ그룹ㆍMDㆍ수량제외를 모두 확정하고, 대시보드는 집계 결과만 표시합니다.

사용법
    python3 build_data.py --src ./원본 --out ../data/latest.json
    python3 build_data.py --src ./원본 --out ../data --split-by-year   # 연도별 파일 분리

필요 패키지
    pip install pandas openpyxl
"""

import argparse, glob, json, os, sys
from datetime import datetime

import pandas as pd
import openpyxl

# ──────────────────────────────────────────────────────────────
# 기초DB 열 위치 (0-based) — 기초DB.xlsx "기초" 시트, 헤더 3행
# ──────────────────────────────────────────────────────────────
COL = {
    "grp_code": 4,    # E  그룹코드 (100/200/300/400/500/600/800)
    "grp_name": 5,    # F  그룹명
    "item_cat": 6,    # G  품목군
    "item_code": 8,   # I  품목코드(원코드)  ← 조인 키
    "item_name": 10,  # K  품목명
    "md_direct": 11,  # L  담당MD (직영)
    "md_b2b": 12,     # M  담당MD (B2B)
    "qty_excl": 15,   # P  수량제외 ("X"면 수량 합산 제외)
    "b2b_corp": 36,   # AK B2B 거래선코드
    "b2b_flag": 37,   # AL "B2B" / "비가전" / "전략" / "직영"
    "corp_code": 41,  # AP 거래선코드
    "corp_name": 42,  # AQ 거래처명
    "st_code": 67,    # BP 지점코드
    "st_name": 68,    # BQ 지점명
    "st_form": 71,    # BT 형태
    "st_class": 72,   # BU 분류
    "org_st_name": 90,  # CM 지점명
    "org_hq": 94,       # CQ 본부
    "org_jibu": 96,     # CS 지사명(지부)
    "org_st_code": 97,  # CT 지점코드
}

GROUP_BY_CODE = {100: "PL", 200: "PS", 300: "PI", 400: "PM",
                 500: "무형", 600: "전략", 800: "물류"}

# ── 채널 판정용 지점코드 ──
STORE_ONLINE = {400040}
STORE_SPECIAL = {302420, 400041}

# ── B2B 판정 ② 삼성전자 + 지정 품목코드 ──
B2B_SAMSUNG_CORP = 1000001
B2B_SAMSUNG_ITEMS = {152, 160, 510, 114, 255, 256, 257, 258}
# ── B2B 판정 ③ 코드 누락 보정 ──
B2B_EXTRA_ITEMS = {268}          # 3D프린터

# ── 지점코드 통합 (지점변경 열로 처리되지 않는 잔여분) ──
STORE_MERGE = {c: 300010 for c in range(300010, 300020)}
STORE_MERGE[300030] = 300010


def load_masters(base_path):
    """기초DB.xlsx에서 품목ㆍ거래선ㆍ지점ㆍB2B거래선 마스터를 읽는다."""
    wb = openpyxl.load_workbook(base_path, read_only=True, data_only=True)
    ws = wb["기초"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))

    def cell(r, key):
        i = COL[key]
        return r[i] if i < len(r) else None

    items, b2b_corps, corps = {}, set(), {}
    st_form, st_org = {}, {}

    for r in rows[3:]:
        # 품목 마스터
        code = cell(r, "item_code")
        if code is not None:
            try:
                code = int(code)
                gc = cell(r, "grp_code")
                items[code] = {
                    "name": cell(r, "item_name"),
                    "grp": GROUP_BY_CODE.get(int(gc)) if isinstance(gc, (int, float)) else None,
                    "md": cell(r, "md_direct"),
                    "md_b2b": cell(r, "md_b2b"),
                    "cat": cell(r, "item_cat"),
                    "qty_excl": str(cell(r, "qty_excl") or "").strip().upper() == "X",
                }
            except (TypeError, ValueError):
                pass

        # B2B 거래선: AL == "B2B" 이고 AK에 "_" 접두가 없는 것만
        # ("_거래선코드" 행은 의도적 표기이므로 제외 — 임의 판단 금지)
        ak, al = cell(r, "b2b_corp"), cell(r, "b2b_flag")
        if al == "B2B" and ak is not None and not str(ak).startswith("_"):
            try:
                b2b_corps.add(int(ak))
            except (TypeError, ValueError):
                pass

        # 거래선 마스터
        cc = cell(r, "corp_code")
        if cc is not None:
            try:
                corps[int(str(cc).lstrip("'"))] = cell(r, "corp_name")
            except (TypeError, ValueError):
                pass

        # 지점 형태/분류 (BP~BU)
        sc = cell(r, "st_code")
        if isinstance(sc, int):
            st_form[sc] = {"name": cell(r, "st_name"),
                           "form": cell(r, "st_form"),
                           "cls": cell(r, "st_class")}

    # 지점 조직 (CM~CT) — 헤더가 한 줄 아래
    for r in rows[4:]:
        oc = r[COL["org_st_code"]] if COL["org_st_code"] < len(r) else None
        jb = r[COL["org_jibu"]] if COL["org_jibu"] < len(r) else None
        if isinstance(oc, int) and jb and jb != "변경지사명":
            st_org.setdefault(oc, {"jibu": jb, "hq": r[COL["org_hq"]]})

    stores = {}
    for code, f in st_form.items():
        o = st_org.get(code, {})
        stores[code] = {"name": f["name"], "form": f["form"], "cls": f["cls"],
                        "jibu": o.get("jibu"), "hq": o.get("hq")}

    print(f"  품목 {len(items):,} / 거래선 {len(corps):,} / 지점 {len(stores):,} / B2B거래선 {len(b2b_corps):,}")
    return items, corps, stores, b2b_corps


def norm_store(row):
    """지점변경 열을 우선 사용하고, 잔여 통합 대상은 추가로 병합한다."""
    code = row["지점변경"] if pd.notna(row["지점변경"]) else row["지점코드"]
    code = int(code)
    return STORE_MERGE.get(code, code)


def classify(df, items, b2b_corps):
    """채널ㆍ그룹ㆍMD를 판정해 컬럼으로 추가한다. (그룹은 B2B를 먼저 확정)"""
    store = df["store"]
    corp = df["거래선코드"].astype("int64")
    item = df["품목코드"].astype("int64")

    # ── STEP 1. B2B 판정 ──
    is_b2b = (
        corp.isin(b2b_corps)
        | ((corp == B2B_SAMSUNG_CORP) & item.isin(B2B_SAMSUNG_ITEMS))
        | item.isin(B2B_EXTRA_ITEMS)
    ).fillna(False)

    # ── STEP 2. 그룹 = B2B 또는 E열 코드 ──
    grp_map = {c: v["grp"] for c, v in items.items()}
    df["grp"] = item.map(grp_map)
    df.loc[is_b2b, "grp"] = "B2B"

    # ── 채널 = 지점코드 우선, 직영 안에서 B2B 분리 ──
    ch = pd.Series("직영", index=df.index, dtype=object)
    ch[store.isin(STORE_ONLINE)] = "온라인"
    ch[store.isin(STORE_SPECIAL)] = "특수"
    ch[(ch == "직영") & is_b2b] = "B2B"
    df["ch"] = ch

    # ── MD = L열, 단 그룹이 B2B면 M열 ──
    md_d = item.map({c: v["md"] for c, v in items.items()})
    md_b = item.map({c: v["md_b2b"] for c, v in items.items()})
    df["md"] = md_d.where(~is_b2b, md_b)

    # ── 수량제외(P열=X): 수량 0 처리 + 단가 산정용 금액(effAmt) 차감 ──
    excl = item.map({c: v["qty_excl"] for c, v in items.items()}).fillna(False).astype(bool)
    df["qty_eff"] = df["수량"].where(~excl, 0)
    df["amt_eff"] = df["금액"].where(~excl, 0)
    return df


def build(src_dir, years=None):
    base = os.path.join(src_dir, "기초DB.xlsx")
    if not os.path.exists(base):
        sys.exit(f"기초DB.xlsx를 찾을 수 없습니다: {base}")

    print("기초DB 로드 중...")
    items, corps, stores, b2b_corps = load_masters(base)

    files = sorted(glob.glob(os.path.join(src_dir, "일보DB_*.csv")))
    if not files:
        sys.exit(f"일보DB_*.csv를 찾을 수 없습니다: {src_dir}")

    facts = []
    for f in files:
        print(f"처리 중: {os.path.basename(f)}")
        df = pd.read_csv(f, encoding="cp949", low_memory=False)
        df = df.dropna(subset=["년월", "지점코드", "품목코드", "거래선코드"])
        # 실데이터에 "(비어 있음)" 등 비정상 값이 섞여 있으므로 숫자 강제 변환 후 제거
        before = len(df)
        for c in ["년월", "지점코드", "품목코드", "거래선코드", "수량", "금액", "지점변경", "판매구분"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        df = df.dropna(subset=["년월", "지점코드", "품목코드", "거래선코드", "판매구분"])
        dropped = before - len(df)
        if dropped:
            print(f"    비정상 코드값 {dropped:,}행 제외")
        if df.empty:
            continue

        df["store"] = df.apply(norm_store, axis=1)
        df["y"] = (df["년월"].astype(int) // 100)
        df["m"] = (df["년월"].astype(int) % 100)
        if years:
            df = df[df["y"].isin(years)]
            if df.empty:
                continue
        df["sub"] = df["금액"].where(df["매입구분"] == "E", 0)   # 매입구분 E = 구독
        df = classify(df, items, b2b_corps)

        g = (df.groupby(["y", "m", "판매구분", "ch", "grp", "md",
                         "품목코드", "거래선코드", "store"], dropna=False)
               .agg(amt=("금액", "sum"), qty=("qty_eff", "sum"),
                    effAmt=("amt_eff", "sum"), sub=("sub", "sum"))
               .reset_index())

        for r in g.itertuples(index=False):
            facts.append((int(r.y), int(r.m), int(r[2]), r.ch, r.grp, r.md,
                          int(r.품목코드), int(r.거래선코드), int(r.store),
                          int(r.amt), int(r.qty), int(r.effAmt), int(r.sub)))
        print(f"    → {len(g):,}행 집계 (누적 {len(facts):,})")

    # ── 컬럼형(dictionary encoding) 인코딩 — 객체 배열 대비 용량 1/4 수준 ──
    ch_d, grp_d, md_d = {}, {}, {}
    def clean(v):
        # pandas 결측치(NaN)는 JSON 표준이 아니므로 반드시 문자열로 정규화한다
        if v is None:
            return ""
        if isinstance(v, float):
            return "" if v != v else str(v)      # NaN 체크
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none") else s

    def idx(d, v):
        v = clean(v)
        if v not in d:
            d[v] = len(d)
        return d[v]

    rows = [[y, m, sd, idx(ch_d, ch), idx(grp_d, grp), idx(md_d, md),
             it, cp, st, amt, qty, eff, sub]
            for (y, m, sd, ch, grp, md, it, cp, st, amt, qty, eff, sub) in facts]

    def keys_of(d):
        return [k for k, _ in sorted(d.items(), key=lambda x: x[1])]

    return {
        "meta": {
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "unit": {"amt": "원(VAT-)", "qty": "대",
                     "note": "단가는 화면에서 ×1.1 하여 VAT+ 표시"},
            "rules": {
                "channel": "온라인=400040 / 특수=302420,400041 / 그 외 직영, 직영 내 그룹=B2B면 B2B",
                "group": "B2B(AL=B2B 거래선 | 1000001+지정품목8 | 268) 우선 확정 후 E열 코드",
                "md": "L열(직영), 그룹이 B2B면 M열",
                "qty": "P열=X 품목은 수량ㆍeffAmt에서 제외",
                "store": "지점변경 열 적용 + 30001*ㆍ300030 → 300010",
                "sd": "1=판매 / 2=매출",
            },
            "years": sorted({r[0] for r in rows}) if rows else [],
        },
        "dims": {"ch": keys_of(ch_d), "grp": keys_of(grp_d), "md": keys_of(md_d)},
        "cols": ["y", "m", "sd", "ch", "grp", "md",
                 "item", "corp", "store", "amt", "qty", "effAmt", "sub"],
        "masters": {
            "items": [{"code": c, "name": clean(v["name"]), "grp": v["grp"],
                       "md": clean(v["md"]), "md_b2b": clean(v.get("md_b2b")),
                       "cat": clean(v.get("cat"))}
                      for c, v in sorted(items.items())],
            "vendors": [{"code": c, "name": clean(n)} for c, n in sorted(corps.items())],
            "stores": [{"code": c, "name": clean(v["name"]), "jibu": clean(v["jibu"]),
                        "form": clean(v["form"]), "cls": clean(v["cls"])}
                       for c, v in sorted(stores.items())],
        },
        "rows": rows,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True, help="일보DB_*.csv 와 기초DB.xlsx 가 있는 폴더")
    ap.add_argument("--out", required=True, help="출력 JSON 경로 (또는 --split-by-year 시 폴더)")
    ap.add_argument("--split-by-year", action="store_true", help="연도별 파일로 분리 저장")
    ap.add_argument("--years", help="처리할 연도 (예: 2024,2025,2026)")
    a = ap.parse_args()

    years = {int(y) for y in a.years.split(",")} if a.years else None
    data = build(a.src, years)

    if a.split_by_year:
        os.makedirs(a.out, exist_ok=True)
        idx = {"meta": data["meta"], "masters": data["masters"], "files": {}}
        for y in data["meta"]["years"]:
            sub = [r for r in data["rows"] if r[0] == y]
            fn = f"facts_{y}.json"
            with open(os.path.join(a.out, fn), "w", encoding="utf-8") as fp:
                json.dump({"y": y, "cols": data["cols"], "dims": data["dims"], "rows": sub},
                      fp, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
            idx["files"][str(y)] = fn
            print(f"저장: {fn} ({len(sub):,}행)")
        with open(os.path.join(a.out, "index.json"), "w", encoding="utf-8") as fp:
            json.dump(idx, fp, ensure_ascii=False, separators=(",", ":"))
        print("저장: index.json")
    else:
        os.makedirs(os.path.dirname(os.path.abspath(a.out)) or ".", exist_ok=True)
        with open(a.out, "w", encoding="utf-8") as fp:
            json.dump(data, fp, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
        mb = os.path.getsize(a.out) / 1024 / 1024
        print(f"저장 완료: {a.out} ({len(data['rows']):,}행, {mb:.1f}MB)")
        if mb > 40:
            print("  ⚠ 파일이 큽니다. --split-by-year 사용을 권장합니다.")


if __name__ == "__main__":
    main()
