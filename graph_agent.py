import os
import json
import re
import openpyxl
from google import genai
from google.genai import types
from dotenv import load_dotenv

load_dotenv()

# --- モデル設定 ---
MODEL_LOW = "gemini-3.1-flash-lite-preview"  # 軽量処理（Excel列番号解析等）

MARKERS = ["*", "x", "o", "^", "square", "+", "diamond"]
LEGEND_LOCATIONS = ["southeast", "northeast", "northwest", "southwest", "best"]

SYSTEM_PROMPT = (
    "あなたはExcelデータを解析してMATLABグラフ生成を支援する専門家です。"
    "ユーザーの指定した列名をデータから探し、JSON形式のみで回答してください。"
    "余分な説明やMarkdownのコードブロックは不要です。"
)

class GraphAgent:

    def __init__(self):
        self.client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

    # --- ① ユーザー入力収集 ---

    def collect_axes(self) -> dict:
        """横軸・縦軸のデータ名をユーザーに入力させる"""
        print("\n--- グラフの軸設定 ---")
        x_name = input("横軸のデータ名を入力してください（例: time）: ").strip()

        y_input = input("縦軸のデータ名を入力してください（複数はカンマ区切り、例: Ta, Tb, Tc）: ")
        y_names = [name.strip() for name in y_input.split(",")]

        x_label = input("横軸のラベル（例: τ: Elapsed Time）: ").strip()
        x_unit  = input("横軸の単位（例: min）: ").strip()
        y_label = input("縦軸のラベル（例: T: Temperature）: ").strip()
        y_unit  = input("縦軸の単位（例: ℃）: ").strip()

        print(f"凡例の位置の選択肢: {LEGEND_LOCATIONS}")
        legend_loc = input("凡例の位置（デフォルト: southeast）: ").strip()
        if legend_loc not in LEGEND_LOCATIONS:
            legend_loc = "southeast"

        y_min = input("Y軸の最小値（空欄で自動）: ").strip()
        y_max = input("Y軸の最大値（空欄で自動）: ").strip()
        ylim = None
        if y_min and y_max:
            ylim = [float(y_min), float(y_max)]

        fig_width  = input("図の幅px（デフォルト: 800）: ").strip() or "800"
        fig_height = input("図の高さpx（デフォルト: 600）: ").strip() or "600"

        return {
            "x_name": x_name,
            "y_names": y_names,
            "x_label": x_label,
            "x_unit": x_unit,
            "y_label": y_label,
            "y_unit": y_unit,
            "legend_location": legend_loc,
            "ylim": ylim,
            "fig_width": int(fig_width),
            "fig_height": int(fig_height),
        }

    # --- ② xlsxをテキスト化してGeminiで解析 ---

    def analyze_xlsx(self, xlsx_path: str, axes: dict) -> dict:
        """
        xlsxをテキスト化してGeminiに渡し、列番号のみを解析させる。
        戻り値例：
        {
            "x_col_idx": 3,
            "y_col_idxs": [4, 5, 6, 7, 8],
        }
        """
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        xlsx_text = ""
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            xlsx_text += f"=== シート: {sheet_name} ===\n"
            # 列番号ヘッダーを先頭行に追加（空列があっても列番号がずれないようにする）
            col_header = "\t".join(f"[col{i}]" for i in range(1, ws.max_column + 1))
            xlsx_text += col_header + "\n"
            for row in ws.iter_rows(values_only=True):
                row_text = "\t".join("" if v is None else str(v) for v in row)
                if row_text.strip():
                    xlsx_text += row_text + "\n"
            xlsx_text += "\n"

        x = axes["x_name"]
        ys = ", ".join(axes["y_names"])

        prompt = (
            "以下のExcelデータを解析してください。\n\n"
            f"{xlsx_text}\n"
            f'探す列：\n- 横軸: "{x}"\n- 縦軸: {ys}\n\n'
            "以下の条件で解析してください：\n"
            "1. 横軸と縦軸が両方含まれているシート・表を特定する\n"
            "2. 同じ名称が複数の表に登場する場合は、横軸と縦軸が共存する表を優先する\n"
            "3. 列番号はExcel全体で左から何番目か（1始まり）を返す\n\n"
            "以下のJSON形式のみで回答してください（説明文不要）：\n"
            "{\n"
            '  "x_col_idx": 横軸の列番号,\n'
            '  "y_col_idxs": [縦軸1の列番号, 縦軸2の列番号, ...]\n'
            "}"
        )

        response = self.client.models.generate_content(
            model=MODEL_LOW,
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_PROMPT,
                max_output_tokens=256,
            )
        )

        text = response.text.strip()
        # コードブロックを除去
        text = re.sub(r"```json|```", "", text).strip()
        # JSONブロックを正規表現で抽出（余分なテキストが混入する場合の保険）
        match = re.search(r'{.*?}', text, re.DOTALL)
        if not match:
            raise ValueError(f"Geminiの返答からJSONを抽出できませんでした: {text}")
        text = match.group(0)
        return json.loads(text)

    # --- ③ .matファイル生成 ---

    def save_mat(self, xlsx_path: str, axes: dict, analysis: dict, mat_filename: str):
        """
        xlsxから該当列を抽出してscipy.io.savematで.matファイルを生成する。
        変数名はユーザー入力（axes）のものを使う。
        """
        import scipy.io
        import numpy as np

        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb.active

        # 数値データ行のみ抽出
        data_rows = [
            row for row in ws.iter_rows(values_only=True)
            if any(isinstance(v, (int, float)) for v in row)
        ]

        x_idx = analysis["x_col_idx"] - 1  # 0始まりに変換
        y_idxs = [i - 1 for i in analysis["y_col_idxs"]]

        mat_data = {}
        mat_data[axes["x_name"]] = np.array(
            [row[x_idx] for row in data_rows], dtype=float
        )
        for y_name, y_idx in zip(axes["y_names"], y_idxs):
            mat_data[y_name] = np.array(
                [row[y_idx] if row[y_idx] is not None else float("nan")
                 for row in data_rows], dtype=float
            )

        scipy.io.savemat(mat_filename, mat_data)
        print(f"{mat_filename} を生成しました。")

    # --- ④ MATLABコード生成（APIなし） ---

    def generate_matlab(self, axes: dict, mat_filename: str) -> str:
        """
        .matを読み込んでグラフを描画するMATLABコードを生成する。
        save_mat()で生成した.matはload一発で変数が展開される。
        """
        x_name = axes["x_name"]
        y_names = axes["y_names"]
        x_label = f"{axes['x_label']} [{axes['x_unit']}]"
        y_label = f"{axes['y_label']} [{axes['y_unit']}]"

        ylim = axes.get("ylim")
        fig_width  = axes.get("fig_width", 800)
        fig_height = axes.get("fig_height", 600)

        lines = []
        lines.append(f"load {mat_filename}")
        lines.append("")
        lines.append(f"fig = figure('Position', [100, 100, {fig_width}, {fig_height}]);")

        lines.append(f'plot({x_name}, {y_names[0]}, "{MARKERS[0]}")')
        lines.append(f'xlabel("{x_label}")')
        lines.append(f'ylabel("{y_label}")')

        if len(y_names) > 1:
            lines.append("hold on")
            for i, y_name in enumerate(y_names[1:], 1):
                lines.append(f'plot({x_name}, {y_name}, "{MARKERS[i % len(MARKERS)]}")')

        legend_labels = ", ".join(f"'{y}'" for y in y_names)
        lines.append(f"legend({legend_labels}, 'Location', '{axes['legend_location']}')")

        if len(y_names) > 1:
            lines.append("hold off")

        if ylim:
            lines.append(f"ylim([{ylim[0]} {ylim[1]}]);")

        lines.append("")
        lines.append("set(fig, 'Color', 'white');")
        lines.append("set(gca, 'Color', 'white');")
        lines.append("set(gca, 'XColor', 'black', 'YColor', 'black');")
        lines.append("leg = legend;")
        lines.append("set(leg, 'Color', 'white', 'TextColor', 'black', 'EdgeColor', 'black');")
        lines.append("saveas(fig, 'graph.png');")
        lines.append("exit;")

        return "\n".join(lines)

    # --- 保存 ---

    def save_matlab(self, code: str, filename: str):
        """生成したMATLABコードを.mファイルに保存する"""
        with open(filename, "w", encoding="utf-8") as f:
            f.write(code)
        print(f"{filename} に保存しました。")
