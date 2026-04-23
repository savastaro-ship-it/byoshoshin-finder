#!/usr/bin/env python3
"""
近畿厚生局の「大阪府・医科・施設基準届出受理医療機関名簿」を取得し、
「病初診」を届け出ている医療機関だけを抽出して data.json を書き出す。

使い方:
    pip install pandas openpyxl requests beautifulsoup4
    python3 build_data.py

出力:
    ../web/data.json  ... Webアプリが読むデータ
    ../web/meta.json  ... 基準日など

厚生局は月1回更新。月初にこのスクリプトを走らせて data.json を更新する。
GitHub Actions で cron 月次実行にすれば完全自動化できる。
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
from bs4 import BeautifulSoup

# -----------------------------------------------------------------------------
# 設定
# -----------------------------------------------------------------------------
INDEX_URL = "https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html"
UA = "Mozilla/5.0 (byoshoshin-finder data builder)"

# 「病初診」を示しそうな列名の候補（年度で微妙に揺れるので広めに拾う）
BYOSHOSHIN_COL_CANDIDATES = ["病初診", "病院初診", "初診"]

OUT_DIR = Path(__file__).resolve().parent.parent / "web"
OUT_DIR.mkdir(parents=True, exist_ok=True)


# -----------------------------------------------------------------------------
# 1. インデックスページから最新の医科ZIPのURLを見つける
# -----------------------------------------------------------------------------
def find_latest_ika_zip_url() -> str:
    """厚生局のページから、最新の医科ZIPのURLを拾う。

    ページには「医科」行に (ZIP) リンクがあり、URLは
        https://kouseikyoku.mhlw.go.jp/kinki/2026.3_sisetukijun_ika.zip
    のような形。年月部分は毎月変わる。
    """
    html = requests.get(INDEX_URL, headers={"User-Agent": UA}, timeout=30).text
    # 医科ZIP URLパターン
    pat = re.compile(
        r'https://kouseikyoku\.mhlw\.go\.jp/kinki/[^"\']*sisetukijun_ika\.zip'
    )
    urls = pat.findall(html)
    if not urls:
        raise RuntimeError("医科ZIPのURLが見つかりませんでした（ページ構造が変わった？）")
    # 複数あった場合は最新（年月最大）を選ぶ
    def key(u: str) -> tuple[int, int]:
        m = re.search(r"/(\d{4})\.(\d{1,2})_", u)
        if not m:
            return (0, 0)
        return (int(m.group(1)), int(m.group(2)))
    urls.sort(key=key, reverse=True)
    return urls[0]


# -----------------------------------------------------------------------------
# 2. ZIPを取得して大阪・医科のExcelだけ開く
# -----------------------------------------------------------------------------
def load_osaka_ika_df(zip_url: str) -> pd.DataFrame:
    print(f"ZIP取得中: {zip_url}", file=sys.stderr)
    resp = requests.get(zip_url, headers={"User-Agent": UA}, timeout=120)
    resp.raise_for_status()
    zf = zipfile.ZipFile(io.BytesIO(resp.content))

    # ZIP内から大阪・医科のxlsxを探す
    osaka_name = None
    for name in zf.namelist():
        # 例: "s2026.4_sisetukijun_ika/s2026.4_sisetukijun_osaka_ika.xlsx"
        if "osaka" in name.lower() and name.lower().endswith(".xlsx"):
            osaka_name = name
            break
    if not osaka_name:
        raise RuntimeError(f"ZIPの中に大阪の医科Excelが見つかりません。中身: {zf.namelist()}")

    print(f"Excel読み込み中: {osaka_name}", file=sys.stderr)
    with zf.open(osaka_name) as f:
        # header=None で全部読み、後でヘッダー行を特定する
        df_all = pd.read_excel(f, sheet_name=0, header=None, dtype=str, engine="openpyxl")
    return df_all


# -----------------------------------------------------------------------------
# 3. ヘッダー行を探して正規のDataFrameに整える
# -----------------------------------------------------------------------------
def normalize_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """近畿厚生局のExcelは先頭数行がタイトル・空行なので、実ヘッダー行を探す。"""
    header_row = None
    for i in range(min(20, len(df_raw))):
        row = df_raw.iloc[i].astype(str).tolist()
        # 「医療機関名称」「医療機関所在地」「医療機関番号」等が含まれる行がヘッダー
        joined = "".join(row)
        if "医療機関名称" in joined and "医療機関所在地" in joined:
            header_row = i
            break
    if header_row is None:
        raise RuntimeError("ヘッダー行（医療機関名称を含む行）が見つかりません")

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() if c is not None else "" for c in df_raw.iloc[header_row].tolist()]
    # 空の行を削除（医療機関名称が空のもの）
    name_col = find_col(df, ["医療機関名称"])
    df = df[df[name_col].notna() & (df[name_col].astype(str).str.strip() != "")]
    return df.reset_index(drop=True)


def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    """列名候補のうち、最初にヒットしたものを返す（部分一致もOK）。"""
    for c in df.columns:
        cs = str(c).strip()
        for cand in candidates:
            if cand in cs:
                return c
    raise KeyError(f"該当列が見つからない。候補={candidates}、列={list(df.columns)}")


# -----------------------------------------------------------------------------
# 4. 「病初診」を届け出ている医療機関を抽出
# -----------------------------------------------------------------------------
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

    # 「病初診」列を見つける
    byo_col = None
    for cand in BYOSHOSHIN_COL_CANDIDATES:
        for c in df.columns:
            if cand == str(c).strip() or str(c).strip().startswith(cand):
                byo_col = c
                break
        if byo_col is not None:
            break
    if byo_col is None:
        # 部分一致で「病初診」「病院初診」を含む列を探す
        for c in df.columns:
            s = str(c)
            if "病初診" in s or "病院初診" in s:
                byo_col = c
                break
    if byo_col is None:
        raise KeyError(
            "「病初診」の列が見つかりません。全列名を確認:\n" +
            "\n".join(f"  - {repr(c)}" for c in df.columns)
        )

    print(f"病初診の列名: {repr(byo_col)}", file=sys.stderr)

    # 値があれば「届出済」と判定（空白/NaN/"-"/○なしなど）
    def is_holder(v) -> bool:
        if v is None:
            return False
        s = str(v).strip()
        if s == "" or s.lower() == "nan" or s == "-" or s == "ー":
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


# -----------------------------------------------------------------------------
# 5. 市区町村を住所から抽出（検索用キー）
# -----------------------------------------------------------------------------
# 大阪府の市区町村を列挙。住所文字列に最初にヒットしたものを「市区町村」として付与する。
OSAKA_CITIES = [
    # 大阪市の区
    "大阪市都島区", "大阪市福島区", "大阪市此花区", "大阪市西区", "大阪市港区",
    "大阪市大正区", "大阪市天王寺区", "大阪市浪速区", "大阪市西淀川区", "大阪市東淀川区",
    "大阪市東成区", "大阪市生野区", "大阪市旭区", "大阪市城東区", "大阪市阿倍野区",
    "大阪市住吉区", "大阪市東住吉区", "大阪市西成区", "大阪市淀川区", "大阪市鶴見区",
    "大阪市住之江区", "大阪市平野区", "大阪市北区", "大阪市中央区",
    # 堺市の区
    "堺市堺区", "堺市中区", "堺市東区", "堺市西区", "堺市南区", "堺市北区", "堺市美原区",
    # 市
    "岸和田市", "豊中市", "池田市", "吹田市", "泉大津市", "高槻市", "貝塚市", "守口市",
    "枚方市", "茨木市", "八尾市", "泉佐野市", "富田林市", "寝屋川市", "河内長野市",
    "松原市", "大東市", "和泉市", "箕面市", "柏原市", "羽曳野市", "門真市", "摂津市",
    "高石市", "藤井寺市", "東大阪市", "泉南市", "四條畷市", "交野市", "大阪狭山市",
    "阪南市",
    # 郡・町村（参考）
    "島本町", "豊能町", "能勢町", "忠岡町", "熊取町", "田尻町", "岬町", "太子町",
    "河南町", "千早赤阪村",
]


def detect_city(addr: str) -> str:
    if not addr:
        return ""
    for c in OSAKA_CITIES:
        if c in addr:
            return c
    # 大阪市・堺市の区が取れなかった場合のフォールバック
    m = re.search(r"(大阪市[^\s〒0-9]{1,3}?区)", addr)
    if m:
        return m.group(1)
    m = re.search(r"(堺市[^\s〒0-9]{1,3}?区)", addr)
    if m:
        return m.group(1)
    m = re.search(r"([^\s〒0-9]{2,6}?(?:市|町|村))", addr)
    if m:
        return m.group(1)
    return ""


# -----------------------------------------------------------------------------
# メイン
# -----------------------------------------------------------------------------
def main():
    zip_url = find_latest_ika_zip_url()
    df_raw = load_osaka_ika_df(zip_url)
    df = normalize_df(df_raw)
    records = extract_byoshoshin(df)

    # 市区町村キーを付与
    for r in records:
        r["city"] = detect_city(r["address"])

    # 出力
    (OUT_DIR / "data.json").write_text(
        json.dumps(records, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    # ZIP URLから年月を抽出してmetaに記録
    m = re.search(r"/(\d{4})\.(\d{1,2})_", zip_url)
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
