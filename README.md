# 大阪・病初診 届出医療機関 検索

近畿厚生局が月1回公表している「保険医療機関の施設基準届出受理状況」から、
大阪府内で **病初診**（医科）を届け出ている医療機関だけを抜き出して、
**地名で検索できるWebアプリ**です。

- 公式一覧ページ: <https://kouseikyoku.mhlw.go.jp/kinki/gyomu/gyomu/hoken_kikan/shitei_jokyo_00004.html>
- このアプリは公式データを元にした非公式の検索ビュー

## 構成

```
.
├── scripts/
│   └── build_data.py      # 厚生局からZIP取得 → Excel解析 → web/data.json出力
├── web/
│   ├── index.html         # 検索UI（静的HTML1枚）
│   ├── data.json          # 抽出結果（build_data.pyが生成）
│   └── meta.json          # データ基準日などのメタ情報
└── .github/workflows/
    └── update-data.yml    # 月1で自動更新＋GitHub Pagesへデプロイ
```

- **データ抽出**: Python (pandas + openpyxl)
- **検索UI**: フレームワーク無しの静的HTML/JS。GitHub Pages にそのまま置ける
- **ホスティング**: GitHub Pages 無料

## セットアップ

### 1. ローカルで試す

```bash
pip install requests pandas openpyxl beautifulsoup4
python3 scripts/build_data.py
# → web/data.json と web/meta.json が生成される

# Webアプリを確認
cd web
python3 -m http.server 8080
# ブラウザで http://localhost:8080/
```

### 2. GitHub Pagesで公開（推奨）

1. このフォルダをGitHubリポジトリにpush
2. リポジトリの **Settings → Pages** で Source を **GitHub Actions** に設定
3. **Settings → Actions → General → Workflow permissions** で
   **Read and write permissions** を有効化
4. Actions タブから `update-data` を手動実行
5. 完了すると `https://<username>.github.io/<repo>/` で公開される
6. 以後は毎月2日 03:00 JST に自動更新

## 使い方

ブラウザで公開URLを開いて、検索窓に地名を入力するだけ。

- 例: `門真` → 門真市内の病初診届出医療機関
- 例: `中央区` → 大阪市中央区（＋他の区も「中央区」は存在しないのでヒットは大阪市中央区）
- 例: `茨木 耳鼻` → 茨木市で名称に「耳鼻」を含む
- スペース区切りでAND検索
- URLに `#q=門真` と付ければ直リンク可能（LINEで共有しやすい）

## 「病初診」って？

**病院初診料の注1の加算**（感染症対策等の施設基準）のこと。
届出済みの病院で初診を受けると、点数が加算される区分。

公式の略称と正式名称の一覧はこちら:
<https://kouseikyoku.mhlw.go.jp/kinki/ryakusyou.pdf>

## メンテ

- 列名や項目が年度改定で変わる可能性あり
- その場合 `build_data.py` の `BYOSHOSHIN_COL_CANDIDATES` や `find_col` を調整
