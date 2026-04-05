from chat_controller import ChatController
from graph_agent import GraphAgent

def main():
    controller = ChatController()
    graph_agent = GraphAgent()

    print("=== レポートサポートエージェント ===")
    print("コマンド: graph / report / exit\n")

    while True:
        command = input("コマンドを入力してください: ").strip().lower()

        if command == "exit":
            print("終了します。")
            break

        elif command == "graph":
            # ファイル指定
            xlsx_path = input("xlsxファイルのパスを入力してください: ").strip()
            if not xlsx_path:
                print("ファイルパスが入力されていません。")
                continue

            mat_filename = input("出力する.matファイル名（例: graph_data.mat）: ").strip() or "graph_data.mat"
            m_filename   = input("出力する.mファイル名（例: output.m）: ").strip() or "output.m"

            # 軸設定（① collect_axes）
            axes = graph_agent.collect_axes()

            # グラフ生成（②③④）
            result = controller.run_graph(xlsx_path, axes, mat_filename, m_filename)

            if result["error"]:
                print(f"エラー: {result['error']}")
                continue

            print(f"\n{m_filename} を生成しました。")

            # MATLAB実行
            run = input("MATLABでグラフを表示しますか？ (y/n): ").strip().lower()
            if run == "y":
                print("MATLAB実行中...")
                matlab_result = controller.run_matlab(m_filename)
                if matlab_result["error"]:
                    print(f"MATLABエラー: {matlab_result['error']}")
                else:
                    print("グラフを表示しました。")

        elif command == "report":
            # ファイル指定
            files_input = input("ファイルパスを入力してください（複数はカンマ区切り）: ")
            file_paths = [f.strip() for f in files_input.split(",")]

            force_ocr = input("スキャンPDFですか？ (y/n): ").strip().lower() == "y"
            theme = input("レポートのテーマを入力してください: ").strip()

            print("\n生成中...")
            result = controller.run_report(file_paths, theme, force_ocr)

            if result["error"]:
                print(f"エラー: {result['error']}")
            else:
                print(f"summary.docx / structure.docx を生成しました。")

        else:
            print("不明なコマンドです。graph / report / exit を入力してください。")

if __name__ == "__main__":
    main()
