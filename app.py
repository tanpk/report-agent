import streamlit as st
import os
import tempfile
import shutil
from pathlib import Path
from chat_controller import ChatController

st.set_page_config(page_title="Report Agent", page_icon="📄", layout="wide")

@st.cache_resource
def get_controller():
    return ChatController()

controller = get_controller()

PROJECTS_DIR = Path("projects")
PROJECTS_DIR.mkdir(exist_ok=True)

# --- サイドバー：プロジェクト管理 ---
with st.sidebar:
    st.title("📁 プロジェクト")

    # 新規プロジェクト作成
    with st.expander("＋ 新規プロジェクト"):
        new_name = st.text_input("プロジェクト名", key="new_project_name")
        if st.button("作成"):
            if new_name.strip():
                (PROJECTS_DIR / new_name.strip()).mkdir(exist_ok=True)
                st.success(f"「{new_name}」を作成しました")
                st.rerun()

    # プロジェクト選択
    projects = sorted([p.name for p in PROJECTS_DIR.iterdir() if p.is_dir()])
    selected = st.selectbox("プロジェクトを選択", ["（未選択）"] + projects, key="selected_project")

    if selected != "（未選択）":
        project_path = PROJECTS_DIR / selected
        files = sorted(project_path.iterdir())
        if files:
            st.markdown("**ファイル一覧**")
            for f in files:
                icon = "📊" if f.suffix == ".xlsx" else "📄" if f.suffix in [".pdf", ".docx", ".txt"] else "🖼️" if f.suffix in [".png", ".jpg"] else "📁"
                st.text(f"{icon} {f.name}")
        else:
            st.caption("ファイルがありません")

# --- プロジェクトパス取得 ---
def get_project_path() -> Path | None:
    sel = st.session_state.get("selected_project", "（未選択）")
    if sel == "（未選択）":
        return None
    return PROJECTS_DIR / sel

# --- メインエリア ---
st.title("📄 Report Agent")
tab_graph, tab_report = st.tabs(["📊 グラフ生成", "📝 レポート作成"])

# ===========================
# グラフ生成タブ
# ===========================
with tab_graph:
    st.subheader("グラフ生成")

    proj = get_project_path()

    # プロジェクト内のxlsxも選択肢に表示
    xlsx_options = []
    if proj:
        xlsx_options = [f.name for f in proj.iterdir() if f.suffix == ".xlsx"]

    use_project_xlsx = False
    if xlsx_options:
        use_project_xlsx = st.checkbox("プロジェクト内のxlsxを使用する")

    if use_project_xlsx:
        selected_xlsx = st.selectbox("xlsxを選択", xlsx_options)
        xlsx_path = str(proj / selected_xlsx)
        xlsx_ready = True
    else:
        uploaded_xlsx = st.file_uploader("xlsxファイルをアップロード", type=["xlsx"], key="graph_xlsx")
        xlsx_path = None
        xlsx_ready = uploaded_xlsx is not None

    if xlsx_ready:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**軸の設定**")
            x_name  = st.text_input("横軸のデータ名", placeholder="例: time")
            y_input = st.text_input("縦軸のデータ名（複数はカンマ区切り）", placeholder="例: Ta, Tb, Tc")
            x_label = st.text_input("横軸ラベル", placeholder="例: τ: Elapsed Time")
            x_unit  = st.text_input("横軸の単位", placeholder="例: min")
            y_label = st.text_input("縦軸ラベル", placeholder="例: T: Temperature")
            y_unit  = st.text_input("縦軸の単位", placeholder="例: ℃")

        with col2:
            st.markdown("**グラフの設定**")
            legend_loc = st.selectbox("凡例の位置", ["southeast", "northeast", "northwest", "southwest", "best"])
            col_ymin, col_ymax = st.columns(2)
            with col_ymin:
                y_min = st.text_input("Y軸最小値", placeholder="空欄で自動")
            with col_ymax:
                y_max = st.text_input("Y軸最大値", placeholder="空欄で自動")
            col_w, col_h = st.columns(2)
            with col_w:
                fig_width  = st.number_input("図の幅(px)", value=800, step=100)
            with col_h:
                fig_height = st.number_input("図の高さ(px)", value=600, step=100)
            mat_filename = st.text_input(".matファイル名", value="graph_data.mat")
            m_filename   = st.text_input(".mファイル名", value="output.m")
            run_matlab   = st.checkbox("MATLABでグラフを生成する", value=True)

        if st.button("グラフ生成", type="primary", key="btn_graph"):
            if not x_name or not y_input:
                st.error("横軸・縦軸のデータ名を入力してください")
            else:
                # 出力先の決定
                out_dir = str(proj) if proj else "."
                mat_out = os.path.join(out_dir, mat_filename)
                m_out   = os.path.join(out_dir, m_filename)
                png_out = os.path.join(out_dir, "graph.png")

                # xlsxを一時ファイルに保存（アップロードの場合）
                tmp_path = None
                if not use_project_xlsx:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                        tmp.write(uploaded_xlsx.read())
                        tmp_path = tmp.name
                    # プロジェクトがあれば保存
                    if proj:
                        shutil.copy(tmp_path, proj / uploaded_xlsx.name)
                    current_xlsx = tmp_path
                else:
                    current_xlsx = xlsx_path

                axes = {
                    "x_name": x_name,
                    "y_names": [y.strip() for y in y_input.split(",")],
                    "x_label": x_label, "x_unit": x_unit,
                    "y_label": y_label, "y_unit": y_unit,
                    "legend_location": legend_loc,
                    "ylim": [float(y_min), float(y_max)] if y_min and y_max else None,
                    "fig_width": int(fig_width), "fig_height": int(fig_height),
                }

                with st.spinner("xlsxを解析中..."):
                    result = controller.run_graph(current_xlsx, axes, mat_out, m_out)

                if tmp_path:
                    os.unlink(tmp_path)

                if result["error"]:
                    st.error(f"エラー: {result['error']}")
                else:
                    st.success(f"{m_filename} を生成しました")
                    with open(m_out, "r", encoding="utf-8") as f:
                        st.download_button("📥 .mファイルをダウンロード", f.read(), file_name=m_filename, mime="text/plain")

                    if run_matlab:
                        with st.spinner("MATLAB実行中..."):
                            matlab_result = controller.run_matlab(m_out)
                        if matlab_result["error"]:
                            st.error(f"MATLABエラー: {matlab_result['error']}")
                        elif os.path.exists(png_out):
                            st.success("グラフを生成しました")
                            st.image(png_out, caption="生成されたグラフ")
                            with open(png_out, "rb") as f:
                                st.download_button("📥 グラフ画像をダウンロード", f.read(), file_name="graph.png", mime="image/png")

# ===========================
# レポート作成タブ
# ===========================
with tab_report:
    st.subheader("レポート作成")

    proj = get_project_path()

    # プロジェクト内ファイルを直接使用するか選択
    use_project_files = False
    project_file_paths = []
    if proj:
        proj_files = [f for f in proj.iterdir() if f.suffix in [".pdf", ".txt", ".docx", ".xlsx", ".png", ".jpg", ".jpeg"]]
        if proj_files:
            use_project_files = st.checkbox("プロジェクト内のファイルを使用する")
            if use_project_files:
                selected_files = st.multiselect(
                    "使用するファイルを選択",
                    [f.name for f in proj_files],
                    default=[f.name for f in proj_files]
                )
                project_file_paths = [str(proj / f) for f in selected_files]

    if not use_project_files:
        uploaded_files = st.file_uploader(
            "資料をアップロード（複数可）",
            type=["pdf", "txt", "docx", "xlsx", "png", "jpg", "jpeg"],
            accept_multiple_files=True, key="report_files"
        )
    else:
        uploaded_files = []

    theme     = st.text_input("レポートのテーマ", placeholder="例: 定常法による熱伝導率の測定")
    force_ocr = st.checkbox("スキャンPDFとして強制OCRする")

    # ②出力内容のチェックボックス
    st.markdown("**出力する内容**")
    output_summary = st.checkbox("要点まとめ（summary.docx）", value=True)
    output_report  = st.checkbox("レポート全文（report.docx）", value=True)

    st.markdown("**出力する章を選択**")
    col1, col2, col3 = st.columns(3)
    with col1:
        ch1 = st.checkbox("1. 目的", value=True)
        ch2 = st.checkbox("2. 原理", value=True)
    with col2:
        ch3 = st.checkbox("3. 実験装置および実験方法", value=True)
        ch4 = st.checkbox("4. 結果", value=True)
    with col3:
        ch5 = st.checkbox("5. 考察", value=True)
        ch6 = st.checkbox("6. まとめ", value=True)

    chapters = []
    if ch1: chapters.append("1.目的")
    if ch2: chapters.append("2.原理")
    if ch3: chapters.append("3.実験装置および実験方法")
    if ch4: chapters.append("4.結果")
    if ch5: chapters.append("5.考察")
    if ch6: chapters.append("6.まとめ")

    st.markdown("**トークン設定**")
    unlimited_tokens = st.checkbox("トークン無制限モード", value=False)
    if unlimited_tokens:
        st.caption("⚠️ 無制限モードはAPIコストが増加する場合があります")
        max_tokens_report = None
    else:
        max_tokens_report = st.number_input(
            "最大出力トークン数", min_value=512, max_value=32768,
            value=10240, step=512,
            help="多いほど詳細な出力が得られますが、APIコストが増加します"
        )

    if st.button("レポート生成", type="primary", key="btn_report"):
        has_files = project_file_paths or uploaded_files
        if not has_files or not theme:
            st.error("ファイルとテーマを入力してください")
        elif not chapters:
            st.error("出力する章を1つ以上選択してください")
        elif not output_summary and not output_report:
            st.error("出力する内容を1つ以上選択してください")
        else:
            # ファイルパスの準備
            tmp_paths = []
            if use_project_files:
                all_paths = project_file_paths
            else:
                for f in uploaded_files:
                    suffix = Path(f.name).suffix
                    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
                        tmp.write(f.read())
                        tmp_paths.append(tmp.name)
                    # プロジェクトがあればファイルを保存
                    if proj:
                        with open(proj / f.name, "wb") as pf:
                            f.seek(0)
                            pf.write(f.read())
                all_paths = tmp_paths

            # 出力先
            out_dir = str(proj) if proj else "."

            with st.spinner("レポートを生成中..."):
                result = controller.run_report(
                    all_paths, theme,
                    force_ocr=force_ocr,
                    chapters=chapters,
                    max_tokens=max_tokens_report,
                    output_summary=output_summary,
                    output_report=output_report,
                    output_dir=out_dir,
                )

            for p in tmp_paths:
                os.unlink(p)

            if result["error"]:
                st.error(f"エラー: {result['error']}")
            else:
                st.success("レポートを生成しました")
                for path, label, fname in [
                    (result["summary"], "要点まとめ", "summary.docx"),
                    (result["structure"], "レポート全文", "report.docx"),
                ]:
                    if path and os.path.exists(path):
                        with open(path, "rb") as f:
                            st.download_button(
                                f"📥 {label}をダウンロード", f.read(),
                                file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
