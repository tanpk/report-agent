# report-agent

AIを活用した大学生向けレポート作成・グラフ生成支援ツール。  
提供した資料（PDF・Word・Excel・画像）をもとに要点まとめとレポート構成案をWord形式で出力し、実験データからMATLABグラフ生成コードを自動作成します。

## 機能

### レポート作成
- PDF・Word・Excel・画像（JPG/PNG）・テキストファイルの読み込み
- スキャンPDF・画像のOCR（Gemini API使用）
- 要点まとめ・レポート構成案の自動生成
- Word（.docx）形式での出力
- OCR結果・APIレスポンスのキャッシュによるコスト削減
- DEBUGモード（APIを呼ばずに動作確認可能）

### グラフ生成
- xlsxファイルをAIで解析し、指定した縦軸・横軸データを自動探索
- MATLABグラフ生成コード（.m）の自動作成
- .matファイルの自動生成（scipy使用）
- MATLABをPythonから直接呼び出してグラフをPNG出力

### チャットUI（ターミナル）
- `graph` / `report` / `exit` コマンドで対話形式で操作
- ロジックとUIが分離された設計（将来のWebUI化に対応）

## セットアップ

### 1. 必要なライブラリのインストール

```bash
pip install google-genai pypdf pdf2image pillow python-docx openpyxl python-dotenv scipy
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

### チャット形式で使う（推奨）

```bash
python chat.py
```

起動後、コマンドを入力します：

```
コマンドを入力してください: graph   # グラフ生成
コマンドを入力してください: report  # レポート作成
コマンドを入力してください: exit    # 終了
```

### レポート作成のみ（`main.py`）

`main.py` の設定を変更して実行します：

```python
# 読み込みたいファイルをリストで指定
file_paths = [
    "資料1.pdf",
    "実験データ.xlsx",
    "図表.png",
]

# スキャンPDFを強制OCRする場合は True
force_ocr = False
```

```bash
python main.py
```

テーマを入力すると `summary.docx`（要点まとめ）と `structure.docx`（レポート構成案）が生成されます。

## DEBUGモード

APIを呼ばずに動作確認したい場合は `agent.py` の先頭を変更します。

```python
DEBUG = True
```

## ファイル構成

```
report-agent/
├── chat.py              # チャット形式のターミナルUI（エントリーポイント）
├── chat_controller.py   # ロジック層（UI非依存）
├── main.py              # レポート作成専用エントリーポイント
├── agent.py             # レポートAIエージェント（生成・キャッシュ管理）
├── graph_agent.py       # グラフ生成エージェント（xlsx解析・MATLABコード生成）
├── file_reader.py       # ファイル読み込み・OCR処理
├── .env                 # APIキー（Gitには含めない）
├── .cache/              # APIレスポンスキャッシュ（自動生成）
└── .ocr_cache/          # OCR結果キャッシュ（自動生成）
```

## 注意事項

- `.env`・`.cache/`・`.ocr_cache/` はGitに含めないようにしてください（`.gitignore`設定済み）
- 使用モデル：`gemini-2.5-flash-lite`（無料枠での利用を想定）
- グラフ生成のxlsx解析にもGemini APIを使用します（列番号の自動特定）
