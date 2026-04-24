#!/usr/bin/env python3
"""
近畿厚生局の「大阪府・医科・施設基準届出受理医療機関名簿」を取得し、
「病初診」を届け出ている医療機関だけを抽出して data.json を書き出す。
"""
from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests
import pandas as pd

# -----------------------------------------------------------------------------
INDEX_URL = "https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html"

# 厚生局は User-Agent 無しアクセスを弾くので、ブラウザっぽいUAを使う
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

BYOSHOSHIN_COL_CANDIDATES = ["病初診", "病院初診", "初診"]
OUT_DIR = Path(__file__).resolve().parent.parent / "web"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def fetch_index_html() -> str:
    r = requests.get(INDEX_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    # レスポンスのエンコーディングを content から推定させる
    r.encoding = r.apparent_encoding or "utf-8"
    return r.text


def find_latest_ika_zip_url() -> str:
    html = fetch_index_html()

    # パターン1: 通常のZIPリンク
    # 例: https://kouseikyoku.mhlw.go.jp/kinki/2026.3_sisetukijun_ika.zip
    patterns = [
        re.compile(r'https://kouseikyoku\.mhlw\.go\.jp/kinki/[^"\'<>]*sisetukijun_ika\.zip'),
        # 念のため s2026.4_... のようなプレフィックス版にも対応
        re.compile(r'https://kouseikyoku\.mhlw\.go\.jp/kinki/s?\d{4}\.\d{1,2}_sisetukijun_ika\.zip'),
    ]
    urls: set[str] = set()
    for pat in patterns:
        for u in pat.findall(html):
            urls.add(u)

    if not urls:
        # デバッグ情報（URL候補があるか）
        sample = [u for u in re.findall(r'https?://[^"\'<>\s]+', html) if "sisetukijun" in u][:10]
        print("デバッグ: sisetukijunを含むURL候補（先頭10件）:", file=sys.stderr)
        for s in sample:
            print(" ", s, file=sys.stderr)
        raise RuntimeError(
            "医科ZIPのURLが見つかりませんでした（ページ構造が変わった？）\n"
            f"HTML長さ: {len(html)} 文字\n"
            f"HTML先頭500文字: {html[:500]!r}"
        )

    def key(u: str) -> tuple[int, int]:
        m = re.search(r"/s?(\d{4})\.(\d{1,2})_", u)
        if not m:
            return (0, 0)
        return (int(m.group(1)), int(m.group(2)))

    best = sorted(urls, key=key, reverse=True)[0]
    return best


def load_osaka_ika_df(zip_url: str) -> pd.DataFrame:
    print(f"ZIP取得中: {zip_url}", file=sys.stderr)
    resp = requests.get(zip_url, headers=HEADERS, timeout=300)
    resp.raise_for_status()
    zf = zipfile.ZipFile(io.BytesIO(resp.content))

    # 大阪医科のxlsxを探す（ZIP直下でもサブフォルダでもOK）
    osaka_name = None
    for name in zf.namelist():
        low = name.lower()
        if "osaka" in low and "ika" in low and low.endswith(".xlsx"):
            osaka_name = name
            break
    if not osaka_name:
        raise RuntimeError(f"ZIP内に大阪医科Excelなし。中身: {zf.namelist()}")

    print(f"Excel読み込み中: {osaka_name}", file=sys.stderr)
    with zf.open(osaka_name) as f:
        df_all = pd.read_excel(f, sheet_name=0, header=None, dtype=str, engine="openpyxl")
    return df_all


def normalize_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    header_row = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].astype(str).tolist()
        joined = "".join(row)
        if "医療機関名称" in joined and ("医療機関所在地" in joined or "所在地" in joined):
            header_row = i
            break
    if header_row is None:
        # 先頭20行を表示してデバッグ
        print("ヘッダー行が見つからない。先頭20行:", file=sys.stderr)
        for i in range(min(20, len(df_raw))):
            print(f"  [{i}]", df_raw.iloc[i].astype(str).tolist()[:8], file=sys.stderr)
        raise RuntimeError("ヘッダー行（医療機関名称を含む行）が見つかりません")

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() if c is not None else "" for c in df_raw.iloc[header_row].tolist()]
    name_col = find_col(df, ["医療機関名称"])
    df = df[df[name_col].notna() & (df[name_col].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)


def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    for c in df.columns:
        cs = str(c).strip()
        for cand in candidates:
            if cand in cs:
                return c
    raise KeyError(f"該当列なし。候補={candidates}、列={list(df.columns)}")


def extract_byoshoshin(df: pd.DataFrame) -> list[dict]:
    name_col = find_col(df, ["医療機関名称"])
    addr_col = find_col(df, ["医療機関所在地", "所在地"])
    tel_col  = find_col(df, ["電話番号"])
    num_col  = find_col(df, ["医療機関番号"])
    bed_col  = None
    try:
        bed_col = find_col(df, ["病床数"])
    except KeyError:
        pass

    byo_col = None
    for cand in BYOSHOSHIN_COL_CANDIDATES:
        for c in df.columns:
            s = str(c).strip()
            if s == cand or s.startswith(cand):
                byo_col = c
                break
        if byo_col is not None:
            break
    if byo_col is None:
        for c in df.columns:
            s = str(c)
            if "病初診" in s or "病院初診" in s:
                byo_col = c
                break
    if byo_col is None:
        raise KeyError(
            "「病初診」の列が見つかりません。全列名:\n" +
            "\n".join(f"  - {repr(c)}" for c in df.columns)
        )

    print(f"病初診の列名: {repr(byo_col)}", file=sys.stderr)

    def is_holder(v) -> bool:
        if v is None:
            return False
        s = str(v).strip()
        if s == "" or s.lower() == "nan" or s in ("-", "ー", "―"):
            return False
        return True

    mask = df[byo_col].apply(is_holder)
    hit = df[mask].copy()
    print(f"病初診の届出医療機関数: {len(hit)}", file=sys.stderr)

    records = []
    for _, r in hit.iterrows():
        records.append({
            "code":    str(r[num_col]).strip() if pd.notna(r[num_col]) else "",
            "name":    str(r[name_col]).strip(),
            "address": str(r[addr_col]).strip() if pd.notna(r[addr_col]) else "",
            "tel":     str(r[tel_col]).strip() if pd.notna(r[tel_col]) else "",
            "beds":    (str(r[bed_col]).strip() if bed_col and pd.notna(r[bed_col]) else ""),
            "byoshoshin_no": str(r[byo_col]).strip(),
        })
    return records


OSAKA_CITIES = [
    "大阪市都島区","大阪市福島区","大阪市此花区","大阪市西区","大阪市港区",
    "大阪市大正区","大阪市天王寺区","大阪市浪速区","大阪市西淀川区","大阪市東淀川区",
    "大阪市東成区","大阪市生野区","大阪市旭区","大阪市城東区","大阪市阿倍野区",
    "大阪市住吉区","大阪市東住吉区","大阪市西成区","大阪市淀川区","大阪市鶴見区",
    "大阪市住之江区","大阪市平野区","大阪市北区","大阪市中央区",
    "堺市堺区","堺市中区","堺市東区","堺市西区","堺市南区","堺市北区","堺市美原区",
    "岸和田市","豊中市","池田市","吹田市","泉大津市","高槻市","貝塚市","守口市",
    "枚方市","茨木市","八尾市","泉佐野市","富田林市","寝屋川市","河内長野市",
    "松原市","大東市","和泉市","箕面市","柏原市","羽曳野市","門真市","摂津市",
    "高石市","藤井寺市","東大阪市","泉南市","四條畷市","交野市","大阪狭山市","阪南市",
    "島本町","豊能町","能勢町","忠岡町","熊取町","田尻町","岬町","太子町",
    "河南町","千早赤阪村",
]


def detect_city(addr: str) -> str:
    if not addr:
        return ""
    for c in OSAKA_CITIES:
        if c in addr:
            return c
    m = re.search(r"(大阪市[^\s〒0-9]{1,3}?区)", addr)
    if m: return m.group(1)
    m = re.search(r"(堺市[^\s〒0-9]{1,3}?区)", addr)
    if m: return m.group(1)
    m = re.search(r"([^\s〒0-9]{2,6}?(?:市|町|村))", addr)
    if m: return m.group(1)
    return ""


def main():
    zip_url = find_latest_ika_zip_url()
    df_raw = load_osaka_ika_df(zip_url)
    df = normalize_df(df_raw)
    records = extract_byoshoshin(df)

    for r in records:
        r["city"] = detect_city(r["address"])

    (OUT_DIR / "data.json").write_text(
        json.dumps(records, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    m = re.search(r"/s?(\d{4})\.(\d{1,2})_", zip_url)
    year_month = f"{m.group(1)}-{int(m.group(2)):02d}" if m else "unknown"

    jst = timezone(timedelta(hours=9))
    meta = {
        "source_url": INDEX_URL,
        "zip_url": zip_url,
        "data_year_month": year_month,
        "generated_at": datetime.now(jst).isoformat(timespec="seconds"),
        "record_count": len(records),
        "prefecture": "大阪府",
        "category": "医科",
        "filter": "病初診 届出医療機関のみ",
    }
    (OUT_DIR / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"✅ 書き出し完了: {OUT_DIR/'data.json'} ({len(records)}件)")
    print(f"✅ メタ情報: {OUT_DIR/'meta.json'}")


if __name__ == "__main__":
    main()
