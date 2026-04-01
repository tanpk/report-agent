from agent import ReportAgent

agent = ReportAgent()

print("=== レポートサポートエージェント ===\n")

# ファイルをリストで複数指定できます
file_paths = [
    "test.txt",
    "tejoho.pdf",
    "tejoho.xlsx",
    "tejoho.png",
]

# スキャンPDF（画像PDF）の場合は force_ocr=True にする
force_ocr = False

# ファイルの読み込みは1回だけ
print("【ファイル読み込み】")
content = agent.load_files(file_paths, force_ocr=force_ocr)

print("\n" + "="*40 + "\n")

theme = input("レポートのテーマを入力してください：")

# 要約と構成案を1回のAPI呼び出しで同時生成
print("\n【生成中...】")
summary, structure = agent.summarize_and_structure(content, theme)

print("\n【要点まとめ】")
print(summary)
agent.save_docx(summary, "summary.docx")

print("\n" + "="*40 + "\n")

print("【レポート構成案】")
print(structure)
agent.save_docx(structure, "structure.docx")
