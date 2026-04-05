import os
import subprocess
from agent import ReportAgent
from graph_agent import GraphAgent

class ChatController:
    """
    ロジック層。ターミナル・WebUI問わず同じメソッドを呼べる。
    UIは chat.py（ターミナル）または将来のWebUIが担当する。
    """

    def __init__(self):
        self.report_agent = ReportAgent()
        self.graph_agent = GraphAgent()

    # --- グラフ生成 ---

    def run_graph(self, xlsx_path: str, axes: dict, mat_filename: str, m_filename: str) -> dict:
        """
        グラフ生成の一連処理を実行する。
        戻り値: {"mat": matファイルパス, "m": mファイルパス, "error": エラー文字列 or None}
        """
        try:
            # ② xlsxを解析して列番号を取得
            print("xlsxを解析中...")
            analysis = self.graph_agent.analyze_xlsx(xlsx_path, axes)
            print(f"解析結果: {analysis}")

            # ③ .matファイルを生成
            self.graph_agent.save_mat(xlsx_path, axes, analysis, mat_filename)

            # ④ MATLABコードを生成・保存
            code = self.graph_agent.generate_matlab(axes, mat_filename)
            self.graph_agent.save_matlab(code, m_filename)

            return {"mat": mat_filename, "m": m_filename, "error": None}

        except Exception as e:
            return {"mat": None, "m": None, "error": str(e)}

    def run_matlab(self, m_filename: str) -> dict:
        """
        MATLABを呼び出して.mファイルを実行する。
        戻り値: {"success": bool, "error": エラー文字列 or None}
        """
        try:
            result = subprocess.run(
                ["matlab", "-batch", f"run('{m_filename}')"],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode != 0:
                return {"success": False, "error": result.stderr}
            return {"success": True, "error": None}
        except subprocess.TimeoutExpired:
            return {"success": False, "error": "MATLABの実行がタイムアウトしました"}
        except FileNotFoundError:
            return {"success": False, "error": "MATLABが見つかりません。PATHを確認してください"}

    # --- レポート作成 ---

    def run_report(self, file_paths: list[str], theme: str, force_ocr: bool = False) -> dict:
        """
        レポート作成の一連処理を実行する。
        戻り値: {"summary": パス, "structure": パス, "error": エラー文字列 or None}
        """
        try:
            content = self.report_agent.load_files(file_paths, force_ocr=force_ocr)
            summary, structure = self.report_agent.summarize_and_structure(content, theme)
            self.report_agent.save_docx(summary, "summary.docx")
            self.report_agent.save_docx(structure, "structure.docx")
            return {"summary": "summary.docx", "structure": "structure.docx", "error": None}
        except Exception as e:
            return {"summary": None, "structure": None, "error": str(e)}
