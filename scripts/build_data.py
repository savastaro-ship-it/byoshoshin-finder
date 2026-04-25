#!/usr/bin/env python3
"""
近畿厚生局の「大阪府・医科」のデータを2種類取得して突合し、
「病初診」を届け出ている医療機関＋開設者・管理者を data.json に書き出す。

データソース:
  A. 施設基準: s{YYYY}.{M}_sisetukijun_ika.zip
     → 受理記号「病初診」の行だけ抽出
  B. コード内容別医療機関一覧: {YYYY}.{M}_kikanzentai_ika.zip
     → 開設者氏名・管理者氏名・指定年月日 etc.

突合キー: 医療機関番号
  A側: "0108313" (ハイフン無し)
  B側: "01-08313" (ハイフン有)
  → ハイフン・空白除去で正規化して結合
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

BASE = "https://kouseikyoku.mhlw.go.jp/kinki"
INDEX_URL = f"{BASE}/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Referer": INDEX_URL,
}

OUT_DIR = Path(__file__).resolve().parent.parent / "docs"
OUT_DIR.mkdir(parents=True, exist_ok=True)


# =============================================================================
# 共通ユーティリティ
# =============================================================================

def norm_code(s) -> str:
    """医療機関番号を正規化。ハイフン・全角ハイフン・空白を全部除去。"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    return re.sub(r"[\s\-－]", "", str(s))


def jst_now():
    return datetime.now(timezone(timedelta(hours=9)))


def candidate_ym_pairs(months_back: int = 6) -> list[tuple[int, int]]:
    today = jst_now()
    y, m = today.year, today.month
    out = []
    for _ in range(months_back):
        out.append((y, m))
        if m == 1: y -= 1; m = 12
        else: m -= 1
    return out


def fetch_zip(urls: list[str], label: str) -> tuple[str, bytes]:
    """候補URLを順に試し、最初に取れたZIPを返す。"""
    print(f"\n=== {label} のZIPを取得 ===", file=sys.stderr)
    errors = []
    for url in urls:
        try:
            print(f"  試行: {url}", file=sys.stderr)
            r = requests.get(url, headers=HEADERS, timeout=300)
            if r.status_code == 200 and len(r.content) > 10000:
                print(f"  ✓ 取得成功 ({len(r.content):,} bytes)", file=sys.stderr)
                return url, r.content
            errors.append(f"{url} -> HTTP {r.status_code}, {len(r.content)} bytes")
        except requests.RequestException as e:
            errors.append(f"{url} -> {e}")
    raise RuntimeError(f"{label} のZIP取得失敗:\n" + "\n".join(errors))


def open_osaka_xlsx(zip_bytes: bytes) -> tuple[str, pd.DataFrame]:
    """ZIP内の大阪×医科のxlsxを開く。"""
    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    target = None
    for name in zf.namelist():
        low = name.lower()
        if "osaka" in low and "ika" in low and low.endswith(".xlsx"):
            target = name
            break
    if not target:
        raise RuntimeError(f"ZIP内に大阪医科Excelなし。中身: {zf.namelist()}")
    print(f"  Excel展開: {target}", file=sys.stderr)
    with zf.open(target) as f:
        df = pd.read_excel(f, sheet_name=0, header=None, dtype=str, engine="openpyxl")
    return target, df


# =============================================================================
# A. 施設基準データ（病初診の行を抽出）
# =============================================================================

def candidate_sisetu_urls() -> list[str]:
    urls = []
    for y, m in candidate_ym_pairs():
        urls.append(f"{BASE}/s{y}.{m}_sisetukijun_ika.zip")
        urls.append(f"{BASE}/{y}.{m}_sisetukijun_ika.zip")
    return urls


def normalize_sisetu_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """施設基準ExcelからヘッダーがあるDataFrameを作る。"""
    header_row = None
    for i in range(min(20, len(df_raw))):
        row = [str(x) if x is not None else "" for x in df_raw.iloc[i].tolist()]
        joined = "".join(row)
        if "医療機関名称" in joined and ("医療機関所在地" in joined or "所在地" in joined):
            header_row = i
            break
    if header_row is None:
        raise RuntimeError("施設基準Excelのヘッダー行が見つかりません")

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


def extract_byoshoshin(df: pd.DataFrame) -> tuple[str, list[dict]]:
    """受理記号が「病初診」の行だけ抽出。"""
    name_col = find_col(df, ["医療機関名称"])
    # 住所カラムは完全一致を優先
    addr_col = None
    for c in df.columns:
        if str(c).strip() == "医療機関所在地（住所）":
            addr_col = c
            break
    if addr_col is None:
        addr_col = find_col(df, ["住所"])
    tel_col     = find_col(df, ["電話番号"])
    num_col     = find_col(df, ["医療機関番号"])
    notice_col  = find_col(df, ["受理届出名称"])
    sign_col    = find_col(df, ["受理記号"])
    recv_no_col = find_col(df, ["受理番号"])
    start_col = None
    try: start_col = find_col(df, ["算定開始年月日"])
    except KeyError: pass
    bed_col = None
    try: bed_col = find_col(df, ["病床数"])
    except KeyError: pass

    sign = df[sign_col].astype(str).str.strip()
    hit = df[sign == "病初診"].copy()

    print(f"  病初診の届出行数: {len(hit)}", file=sys.stderr)
    print(f"  ユニーク医療機関数: {hit[num_col].astype(str).nunique()}", file=sys.stderr)

    records = []
    seen = set()
    for _, r in hit.iterrows():
        code_raw = str(r[num_col]).strip() if pd.notna(r[num_col]) else ""
        if not code_raw or code_raw in seen:
            continue
        seen.add(code_raw)
        records.append({
            "code":          code_raw,
            "code_norm":     norm_code(code_raw),  # 突合用
            "name":          str(r[name_col]).strip() if pd.notna(r[name_col]) else "",
            "address":       str(r[addr_col]).strip() if pd.notna(r[addr_col]) else "",
            "tel":           str(r[tel_col]).strip() if pd.notna(r[tel_col]) else "",
            "beds":          (str(r[bed_col]).strip() if bed_col and pd.notna(r[bed_col]) else ""),
            "notice_name":   str(r[notice_col]).strip() if pd.notna(r[notice_col]) else "",
            "byoshoshin_no": str(r[recv_no_col]).strip() if pd.notna(r[recv_no_col]) else "",
            "start_date":    (str(r[start_col]).strip() if start_col and pd.notna(r[start_col]) else ""),
        })
    return sign_col, records


# =============================================================================
# B. コード内容別医療機関一覧（開設者・管理者）
# =============================================================================

def candidate_zentai_urls() -> list[str]:
    """コード内容別医療機関一覧のZIP候補URL。
    既知パターン: 2026.4_kikanzentai_ika.zip
    念のため s プレフィックス版にも対応。
    """
    urls = []
    for y, m in candidate_ym_pairs():
        urls.append(f"{BASE}/{y}.{m}_kikanzentai_ika.zip")
        urls.append(f"{BASE}/s{y}.{m}_kikanzentai_ika.zip")
    return urls


def build_kanri_map(df_raw: pd.DataFrame) -> dict[str, dict]:
    """コード内容別Excelから {医療機関番号(正規化): {開設者, 管理者, 指定日, 種別}} の辞書を作る。

    このExcelは1医療機関＝複数行（PDF印刷の見た目をそのままExcel化したもの）なので、
    「項番（A列）が数字の行 = 1医療機関の1行目」だけ拾えばOK。
    カラムは固定インデックス：
      0=項番  1=医療機関番号  2=名称  3=住所  4=電話
      5=開設者氏名  6=管理者氏名  7=指定年月日  8=病床/診療科  9=種別
    """
    def is_main_row(v) -> bool:
        if pd.isna(v): return False
        return str(v).strip().isdigit()

    main = df_raw[df_raw.iloc[:, 0].apply(is_main_row)].copy()
    print(f"  全体一覧のメイン行数: {len(main)}", file=sys.stderr)

    out: dict[str, dict] = {}
    for _, r in main.iterrows():
        code_raw = str(r[1]).strip() if pd.notna(r[1]) else ""
        key = norm_code(code_raw)
        if not key:
            continue
        out[key] = {
            "founder":      _clean(r[5]),
            "manager":      _clean(r[6]),
            "designated":   _clean(r[7]),
            "kind":         _clean(r[9]),
            "code_full":    code_raw,
        }
    print(f"  突合用キー数: {len(out)}", file=sys.stderr)
    return out


def _clean(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    # 全角スペースが多用されてるので半角に寄せて整形（ただし氏名内のスペースは保持）
    s = str(v).strip()
    return s


# =============================================================================
# 市区町村判定
# =============================================================================

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
    if not addr: return ""
    for c in OSAKA_CITIES:
        if c in addr: return c
    m = re.search(r"(大阪市[^\s〒0-9]{1,3}?区)", addr)
    if m: return m.group(1)
    m = re.search(r"(堺市[^\s〒0-9]{1,3}?区)", addr)
    if m: return m.group(1)
    m = re.search(r"([^\s〒0-9]{2,6}?(?:市|町|村))", addr)
    if m: return m.group(1)
    return ""


# =============================================================================
# メイン
# =============================================================================

def main():
    # A. 施設基準
    sisetu_url, sisetu_zip = fetch_zip(candidate_sisetu_urls(), "施設基準")
    sisetu_xlsx, df_sisetu_raw = open_osaka_xlsx(sisetu_zip)
    df_sisetu = normalize_sisetu_df(df_sisetu_raw)
    sign_col, records = extract_byoshoshin(df_sisetu)

    # B. コード内容別医療機関一覧
    zentai_url, zentai_zip = fetch_zip(candidate_zentai_urls(), "全体（コード内容別）")
    zentai_xlsx, df_zentai_raw = open_osaka_xlsx(zentai_zip)
    kanri_map = build_kanri_map(df_zentai_raw)

    # 突合
    matched = 0
    for r in records:
        info = kanri_map.get(r["code_norm"])
        if info:
            r["founder"]    = info["founder"]
            r["manager"]    = info["manager"]
            r["designated"] = info["designated"]
            r["kind"]       = info["kind"]
            matched += 1
        else:
            r["founder"]    = ""
            r["manager"]    = ""
            r["designated"] = ""
            r["kind"]       = ""
        r["city"] = detect_city(r["address"])
        # 突合用フィールドはJSONからは外してサイズ削減
        del r["code_norm"]

    print(f"\n=== 突合結果 ===", file=sys.stderr)
    print(f"  病初診 records: {len(records)}件", file=sys.stderr)
    print(f"  開設者・管理者を取得: {matched}件", file=sys.stderr)
    print(f"  突合できなかった: {len(records) - matched}件", file=sys.stderr)

    # 書き出し
    (OUT_DIR / "data.json").write_text(
        json.dumps(records, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    m1 = re.search(r"/s?(\d{4})\.(\d{1,2})_", sisetu_url)
    sisetu_ym = f"{m1.group(1)}-{int(m1.group(2)):02d}" if m1 else "unknown"
    m2 = re.search(r"/s?(\d{4})\.(\d{1,2})_", zentai_url)
    zentai_ym = f"{m2.group(1)}-{int(m2.group(2)):02d}" if m2 else "unknown"

    meta = {
        "source_url": INDEX_URL,
        "sisetu_zip_url": sisetu_url,
        "zentai_zip_url": zentai_url,
        "sisetu_xlsx": sisetu_xlsx,
        "zentai_xlsx": zentai_xlsx,
        "sisetu_year_month": sisetu_ym,
        "zentai_year_month": zentai_ym,
        "data_year_month": sisetu_ym,  # 表示用は施設基準側を採用
        "generated_at": jst_now().isoformat(timespec="seconds"),
        "record_count": len(records),
        "matched_kanri": matched,
        "prefecture": "大阪府",
        "category": "医科",
        "filter": "病初診 届出医療機関のみ＋開設者・管理者突合",
    }
    (OUT_DIR / "meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\n✅ data.json: {len(records)}件")
    print(f"✅ meta.json: 開設者・管理者突合 {matched}/{len(records)}")


if __name__ == "__main__":
    main()
