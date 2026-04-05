# report-agent

AIを活用した大学生向けレポート作成・グラフ生成支援ツール。  
実験データや資料をもとに、要点まとめ・レポート全文の自動生成と、MATLABグラフ生成コードの出力をWebUIから行えます。

## 機能

### レポート作成
- PDF・Word・Excel・画像（JPG/PNG）・テキストファイルの読み込み
- スキャンPDF・画像のOCR（Gemini API使用）
- 要点まとめ・レポート全文の自動生成
- Word（.docx）形式での出力
- 出力する章の個別選択
- OCR結果・APIレスポンスのキャッシュによるコスト削減

### グラフ生成
- xlsxファイルをAIで解析し、指定した縦軸・横軸データを自動探索
- MATLABグラフ生成コード（.m）の自動作成
- .matファイルの自動生成
- MATLABを直接呼び出してグラフをPNG出力

### プロジェクト管理
- プロジェクトフォルダを作成してファイルを一元管理
- アップロードしたファイルをプロジェクトに自動保存
- 出力ファイル（docx・mat・m・png）もプロジェクトフォルダに保存

## セットアップ

### 1. 必要なライブラリのインストール

```bash
pip install google-genai pypdf pdf2image pillow python-docx openpyxl python-dotenv scipy streamlit
```

WindowsでスキャンPDFを使う場合は [Poppler](https://github.com/oschwartz10612/poppler-windows/releases) も別途インストールしてください。

グラフ生成機能を使う場合はMATLABがインストールされ、`matlab`コマンドにPATHが通っている必要があります。

### 2. APIキーの設定

プロジェクトルートに `.env` ファイルを作成し、Gemini APIキーを記述します。

```
GEMINI_API_KEY=your_api_key_here
```

APIキーは [Google AI Studio](https://aistudio.google.com/) で取得できます。

## 使い方

### WebUI（推奨）

```bash
streamlit run app.py
```

ブラウザが自動で開きます（`http://localhost:8501`）。

### ターミナル

```bash
python chat.py
```

コマンド：`graph`（グラフ生成）/ `report`（レポート作成）/ `exit`（終了）

## ファイル構成

```
report-agent/
├── app.py               # WebUI（Streamlit）エントリーポイント
├── chat.py              # ターミナルUIエントリーポイント
├── chat_controller.py   # ロジック層（UI非依存）
├── agent.py             # レポートAIエージェント
├── graph_agent.py       # グラフ生成エージェント
├── file_reader.py       # ファイル読み込み・OCR処理
├── main.py              # レポート作成専用スクリプト
├── .env                 # APIキー（Gitには含めない）
├── projects/            # プロジェクトフォルダ（自動生成）
├── .cache/              # APIレスポンスキャッシュ（自動生成）
└── .ocr_cache/          # OCR結果キャッシュ（自動生成）
```

## モデル設定

| 用途 | モデル |
|---|---|
| レポート生成（要点まとめ・全文） | `gemini-3-flash-preview` |
| Excel列番号解析・OCR | `gemini-2.5-flash-lite` |

## 注意事項

- `.env`・`.cache/`・`.ocr_cache/`・`projects/` はGitに含めないようにしてください（`.gitignore`設定済み）
- 無料枠での利用を想定していますが、レポート生成には`gemini-3-flash-preview`を使用するためレート制限に注意してください
