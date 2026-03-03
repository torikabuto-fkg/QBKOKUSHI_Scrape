# QB Scraper

**medilink-study.com QB（クエスチョン・バンク）** の問題・解説を自動スクレイピングし、  
**PDF / Excel / JSON / 画像ファイル** として一括出力する Node.js スクリプトです。

---

## ✨ 特徴

| 機能 | 説明 |
|------|------|
| **4 形式同時出力** | PDF（フル解説付き）、Excel（`.xlsx`）、JSON、画像フォルダ |
| **連問（シリアル問題）対応** | 共通問題文＋枝問をまとめて 1 セットとして処理 |
| **解説セクション完全取得** | 解法の要点 / 選択肢解説 / ガイドライン / 診断 / KEYWORD / 主要所見 / 画像診断 |
| **基本事項セクション** | 「すべて表示」自動クリック → テキスト＋テーブルスクリーンショット取得 |
| **医ンプット（iframe）** | iframe 内のテキスト・画像も自動スクレイピング |
| **画像バイナリ検証** | `image-size` による PNG/JPEG 判定。壊れた画像は自動スキップ |
| **ループ検知** | 同じ問題番号が連続した場合に自動停止（無限ループ防止） |
| **エラー耐性** | 個別問題でエラーが出ても残りの出力を継続 |
| **ポータブル実行** | `module.paths` 自動設定でどのディレクトリからでも実行可能 |

---

## 📋 前提条件

- **Node.js v18 以上**（v22 推奨）  
  - `?.`（Optional Chaining）構文を使用するため、古いバージョンでは動作しません
  - Windows の場合は **WSL (Windows Subsystem for Linux)** での実行を推奨
- **npm**（Node.js 付属）
- **medilink-study.com のアカウント**（QB のログイン情報が必要）

---

## 🚀 セットアップ

### 1. リポジトリをクローン

```bash
git clone https://github.com/<your-username>/qb-scraper.git
cd qb-scraper
```

### 2. 依存パッケージをインストール

```bash
npm install puppeteer pdfmake exceljs axios image-size
```

### 3. 日本語フォントを配置

`fonts/` フォルダを作成し、以下のフォントファイルを配置してください：

```
fonts/
├── NotoSansJP-Regular.ttf
└── NotoSansJP-Bold.ttf
```

📥 ダウンロード: [Google Fonts - Noto Sans JP](https://fonts.google.com/noto/specimen/Noto+Sans+JP)

### 4. 設定ファイルを編集

`QB_Scrape_Ver.5.js` 内の `CONFIG` ブロックを自分の環境に合わせて書き換えます：

```javascript
const CONFIG = {
  loginUrl: 'https://login.medilink-study.com/login',
  email: 'your-email@example.com',       // ← ログイン用メールアドレス
  password: 'your-password',             // ← ログイン用パスワード
  startUrl: 'https://qb.medilink-study.com/Answer/117A1',  // ← 最初の問題ページURL
  numberOfPages: 100,                    // ← 取得する問題数
  fileName: 'A 消化器',                  // ← 出力フォルダ名 & ファイル名（拡張子不要）
};
```

> ⚠️ **注意**: `email` と `password` は **絶対に Git にコミットしないでください**。  
> `.gitignore` でスクリプトごと除外するか、環境変数を使うことを推奨します。

---

## ▶️ 実行

```bash
node QB_Scrape_Ver.5.js
```

### WSL（Windows）の場合

```bash
# デスクトップから直接実行する例
cd /mnt/c/Users/<username>/Desktop
node QB_Scrape_Ver.5.js
```

### 実行ログの例

```
🚀 QB スクレイピング Ver.5
   URL   : https://qb.medilink-study.com/Answer/117A1
   問題数: 100
   出力名: A 消化器

📝 ログイン中...
  ✓ ログイン完了
  ⏳ ページ描画待ち...
  ✓ ページ描画完了
  [1/100] 117A1 ... ✓
  [2/100] 117A2 ... ✓
  ...

📊 スクレイピング完了: 100 問

🖼️  画像保存中... → A 消化器/A 消化器_images/
📄 出力ファイル生成中...
  ✅ PDF: A 消化器/A 消化器.pdf
  ✅ Excel: A 消化器/A 消化器.xlsx
  ✅ JSON: A 消化器/A 消化器.json

🏁 完了
```

---

## 📁 出力ファイル構成

```
A 消化器/
├── A 消化器.pdf          # フル解説付き PDF（問題→解説の見開き構成）
├── A 消化器.xlsx         # Excel 一覧表（フィルター・ウィンドウ枠固定付き）
├── A 消化器.json         # 構造化 JSON データ
└── A 消化器_images/      # 全画像ファイル
    ├── 117A1_問題_1.png
    ├── 117A1_解説_1.png
    ├── 117A1_基本事項_1.png
    ├── 117A1_医ンプット_1.png
    └── ...
```

### 各出力形式の詳細

#### PDF
- 問題ページ → 解説ページ の構成で改ページ
- 画像は 2 列レイアウトで自動配置
- セクション: 正解・正答率 / 主要所見 / KEYWORD / 画像診断 / 診断 / 解法の要点 / 選択肢解説 / ガイドライン / 基本事項 / 医ンプット

#### Excel
- ヘッダー行固定・オートフィルター付き
- 列: 問題番号 / 掲載頁 / 問題文 / 選択肢 / 正解 / 正答率 / 各解説セクション / 画像ファイルパス

#### JSON
- 全データを構造化して保存（画像は Base64 サイズ表記に変換）
- 他のプログラムでの再利用・分析に最適

---

## ⚙️ カスタマイズ

### タイミング設定

`TIMING` 定数でスクレイピングの待機時間を調整できます：

```javascript
const TIMING = {
  initialWait: 8000,       // ページ遷移後の待機 (ms)
  scrollStep: 100,         // スクロール刻み (px)
  scrollInterval: 80,      // スクロール間隔 (ms)
  retryInterval: 500,      // 問題文取得リトライ間隔 (ms)
  maxRetries: 60,          // 問題文取得の最大リトライ数
  afterClick: 2000,        // ボタンクリック後の待機 (ms)
  selectorTimeout: 10000,  // セレクタ待機のタイムアウト (ms)
};
```

> 回線速度やサーバー応答が遅い場合は `initialWait` や `afterClick` を増やしてください。

### PDF レイアウト

```javascript
const PDF_LAYOUT = {
  availableWidth: 515.28,  // A4 横幅 (pt)
  maxImageCols: 3,         // 画像の最大列数
};
```

---

## 🛠️ 技術スタック

| パッケージ | 用途 |
|-----------|------|
| [puppeteer](https://pptr.dev/) | ヘッドレス Chrome による Web スクレイピング |
| [pdfmake](http://pdfmake.org/) | PDF 生成（日本語フォント対応） |
| [exceljs](https://github.com/exceljs/exceljs) | Excel (.xlsx) 生成 |
| [axios](https://axios-http.com/) | 画像ダウンロード |
| [image-size](https://github.com/image-size/image-size) | 画像フォーマット・サイズ検証 |

---

## 📌 注意事項

- このスクリプトは **個人の学習目的** で使用してください
- 短時間に大量リクエストを送るとサーバーに負荷がかかります。`TIMING` の値を適切に設定してください
- **ログイン情報（email / password）を GitHub に公開しないでください**
- スクレイピング対象サイトの利用規約を確認の上ご利用ください

---

## 📝 更新履歴

### Ver.5（最新）
- 実 HTML の DOM 構造に完全対応（セレクタ修正）
- 解答ボタン: `#answerSection` / `#answerCbtSection` 両対応
- 「医ンプット」セクションの iframe 内スクレイピングに対応
- 「基本事項」の「すべて表示」自動クリック
- Excel (`.xlsx`) 出力を追加
- JSON 出力を追加
- 画像ファイル個別保存を追加
- `image-size` によるバイナリ検証で壊れた画像をスキップ
- `node-fetch` 不要（Node.js v22 組み込み `fetch` 使用）
- `PdfPrinter` で正規フォント読み込み（`vfs_fonts.js` 不要）
- ポータブル実行対応（`module.paths` 自動設定）

---

## License

MIT
