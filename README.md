# report-agent

AIを活用した大学生向けレポート作成支援ツール。  
提供した資料（PDF・Word・Excel・画像）のみをもとに、要点まとめとレポート構成案をWord形式で出力します。

## 機能

- PDF・Word・Excel・画像（JPG/PNG）・テキストファイルの読み込み
- スキャンPDF・画像のOCR（Gemini API使用）
- 要点まとめ・レポート構成案の自動生成
- Word（.docx）形式での出力
- OCR結果・APIレスポンスのキャッシュによるコスト削減
- DEBUGモード（APIを呼ばずに動作確認可能）

## セットアップ

### 1. 必要なライブラリのインストール

```bash
pip install google-genai pypdf pdf2image pillow python-docx openpyxl python-dotenv
```

Windowsでスキャン PDF を使う場合は [Poppler](https://github.com/oschwartz10612/poppler-windows/releases) も別途インストールしてください。

### 2. APIキーの設定

プロジェクトルートに `.env` ファイルを作成し、Gemini APIキーを記述します。

```
GEMINI_API_KEY=your_api_key_here
```

APIキーは [Google AI Studio](https://aistudio.google.com/) で取得できます。

## 使い方

### 1. `main.py` の設定

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

### 2. 実行

```bash
python main.py
```

テーマを入力すると、`summary.docx`（要点まとめ）と `structure.docx`（レポート構成案）が生成されます。

## DEBUGモード

APIを呼ばずに動作確認したい場合は `agent.py` の先頭を変更します。

```python
DEBUG = True
```

## ファイル構成

```
report-agent/
├── main.py          # エントリーポイント
├── agent.py         # AIエージェント（生成・キャッシュ管理）
├── file_reader.py   # ファイル読み込み・OCR処理
├── .env             # APIキー（Gitには含めない）
├── .cache/          # APIレスポンスキャッシュ（自動生成）
└── .ocr_cache/      # OCR結果キャッシュ（自動生成）
```

## 注意事項

- `.env`・`.cache/`・`.ocr_cache/` はGitに含めないようにしてください（`.gitignore`設定済み）
- 使用モデル：`gemini-2.5-flash-lite`（無料枠での利用を想定）
