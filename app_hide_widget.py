import streamlit as st
import os
import tempfile
import base64
from datetime import datetime
from report_logic import generate_report
from report_logic_advanced import generate_report as generate_report_advanced

# ─────────────────────────────────────────────────────
# アイコン画像パス
# ─────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_TITLE = os.path.join(SCRIPT_DIR, "image1.png")
ICON_SETTING = os.path.join(SCRIPT_DIR, "image2.png")
ICON_PDF = os.path.join(SCRIPT_DIR, "iamge3.png")

# ─────────────────────────────────────────────────────
# ユーティリティ
# ─────────────────────────────────────────────────────
def icon_b64(path):
    """画像ファイルをbase64文字列に変換"""
    if os.path.exists(path):
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode("utf-8")
    return None

def icon_img_tag(path, size=28):
    """インラインで使えるimg HTMLタグを返す"""
    b64 = icon_b64(path)
    if b64:
        return f'<img src="data:image/png;base64,{b64}" width="{size}" style="vertical-align: middle; margin-right: 8px;">'
    return ""

def display_pdf(pdf_bytes):
    """PDFをプレビュー表示する"""
    from streamlit_pdf_viewer import pdf_viewer
    pdf_viewer(pdf_bytes, width=700)

# ─────────────────────────────────────────────────────
# Streamlit UI 設定
# ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="GA4 Report Generator",
    page_icon="📊",
    layout="wide"
)

# カスタムCSS（MOCALカラー + ファイルアップローダー日本語化）
st.markdown("""
    <style>
    .main {
        background-color: #F5F8FC;
    }
    /* Streamlitのフッター・メニュー・Cloud管理UIを非表示 */
    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    header[data-testid="stHeader"] {visibility: hidden;}
    .viewerBadge_container__r5tak {display: none !important;}
    .stActionButton {display: none !important;}
    [data-testid="manage-app-button"] {display: none !important;}
    .styles_viewerBadge__CvC9N {display: none !important;}
    ._profileContainer_gzau3_53 {display: none !important;}
    ._profilePreview_gzau3_63 {display: none !important;}
    [data-testid="stStatusWidget"] {display: none !important;}
    div[data-testid="stStatusWidget"] {display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    #stDecoration {display: none !important;}
    /* 右下の外部リンクバッジ（GitHub / Streamlit）を強制的に無効化 */
    [data-testid="stDecoration"] a[href*="streamlit.io"] {display: none !important;}
    [data-testid="stStatusWidget"] a[href*="streamlit.io"] {display: none !important;}
    [data-testid="stStatusWidget"] a[href*="github.com"] {display: none !important;}
    [data-testid="stStatusWidget"] a {display: none !important;}
    [data-testid="stStatusWidget"] * {pointer-events: none !important;}
    .st-emotion-cache-h4xjwg {display: none !important;}
    .ea3mdgi5 {display: none !important;}
    div[class*="stDeployButton"] {display: none !important;}
    div[class*="StatusWidget"] {display: none !important;}
    button[kind="manage"] {display: none !important;}
    [class*="viewerBadge"] {display: none !important;}
    [class*="statusWidget"] {display: none !important;}
    [class*="StatusWidget"] {display: none !important;}
    /* プレースホルダーをより薄いグレーに */
    input::placeholder {color: #C8D0D8 !important; opacity: 1 !important;}
    .stButton>button {
        background-color: #1A6FAB;
        color: white;
        border-radius: 5px;
        width: 100%;
        height: 3em;
        font-weight: bold;
    }
    .stDownloadButton>button {
        background-color: #4DA3D4;
        color: white;
        border-radius: 5px;
        width: 100%;
        height: 3em;
        font-weight: bold;
    }
    /* ファイルアップローダーの英語テキストを日本語に置換 */
    [data-testid="stFileUploaderDropzone"] div:has(> small) div:first-child {
        visibility: hidden;
        position: relative;
        height: 1.2em;
    }
    [data-testid="stFileUploaderDropzone"] div:has(> small) div:first-child::after {
        content: "ここにファイルをドラッグ＆ドロップ";
        visibility: visible;
        position: absolute;
        top: 0;
        left: 0;
    }
    [data-testid="stFileUploaderDropzone"] small {
        visibility: hidden;
        position: relative;
        height: 1.2em;
        display: inline-block;
    }
    [data-testid="stFileUploaderDropzone"] small::after {
        content: "ファイルサイズ上限: 200MB";
        visibility: visible;
        position: absolute;
        top: 0;
        left: 0;
        white-space: nowrap;
    }
    [data-testid="stFileUploaderDropzone"] button {
        visibility: hidden;
        position: relative;
    }
    [data-testid="stFileUploaderDropzone"] button::after {
        content: "ファイルを選択";
        visibility: visible;
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        white-space: nowrap;
    }
    </style>
    """, unsafe_allow_html=True)

# ヘッダー
_title_icon = icon_img_tag(ICON_TITLE, size=40)
st.markdown(
    f'<h1 style="display:flex; align-items:center;">{_title_icon}GA4 月次レポート自動生成ツール</h1>',
    unsafe_allow_html=True
)
st.write("GA4からエクスポートしたCSVをアップロードするだけで、整形されたPDFレポートを作成します。")

st.divider()

# サイドバー設定
_setting_icon = icon_img_tag(ICON_SETTING, size=28)
st.sidebar.markdown(
    f'<h2 style="display:flex; align-items:center;">{_setting_icon}レポート設定</h2>',
    unsafe_allow_html=True
)
company_name = st.sidebar.text_input("会社名", value="", placeholder="入力例）ABC株式会社")
staff_name = st.sidebar.text_input("担当者名", value="", placeholder="入力例）サンプル 太郎")

st.sidebar.divider()
report_type = st.sidebar.radio(
    "レポートの種類",
    ["スタンダード (5ページ)", "詳細版 (10ページ)"],
    index=1,
    help="【スタンダード】数値の簡易報告向け。コメントなしで主要KPI・チャネル・CV・上位ページをシンプルに表示します。\n\n【詳細版】分析・改善提案向け。成果CV/補助CVの分離、CVR分析、フォーム離脱分析、ユーザーリテンション、コンテンツ・曜日分析、Paid Social分析、計測定義・注記を含む10ページ構成です。"
)

# メインコンテンツ
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. データのインポート")
    uploaded_csv = st.file_uploader("GA4エクスポートCSVを選択してください", type=["csv"])

    uploaded_logo = st.file_uploader("ロゴを入れたい場合（任意）", type=["png", "jpg", "jpeg"])


with col2:
    st.subheader("2. レポートの生成")
    if uploaded_csv is not None:
        st.success("CSVファイルが読み込めました！")

        # session_stateでPDF生成結果を保持（ボタン押下後も消えないように）
        if st.button("🚀 PDFレポートを生成する"):
            with st.spinner("レポートを生成中です。しばらくお待ちください..."):
                try:
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        csv_path = os.path.join(tmp_dir, "input.csv")
                        with open(csv_path, "wb") as f:
                            f.write(uploaded_csv.getbuffer())

                        logo_path = None
                        if uploaded_logo:
                            logo_path = os.path.join(tmp_dir, "logo.png")
                            with open(logo_path, "wb") as f:
                                f.write(uploaded_logo.getbuffer())

                        output_pdf_path = os.path.join(tmp_dir, "report.pdf")

                        if "スタンダード" in report_type:
                            generate_report(
                                csv_path,
                                output_pdf_path,
                                company_name=company_name,
                                department=staff_name,
                                logo_path=logo_path
                            )
                        else:
                            generate_report_advanced(
                                csv_path,
                                output_pdf_path,
                                company_name=company_name,
                                department=staff_name,
                                logo_path=logo_path
                            )

                        with open(output_pdf_path, "rb") as pdf_file:
                            st.session_state['pdf_data'] = pdf_file.read()

                    st.session_state['pdf_ready'] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

        # 生成済みPDFがあれば表示
        if st.session_state.get('pdf_ready'):
            pdf_data = st.session_state['pdf_data']
            st.success("レポートの生成が完了しました！")

            st.download_button(
                label="📥 PDFをダウンロードする",
                data=pdf_data,
                file_name=f"GA4レポート_{company_name}_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )
    else:
        st.info("CSVファイルをアップロードしてください。")
        # CSVが外されたらPDF状態もクリア
        if 'pdf_ready' in st.session_state:
            del st.session_state['pdf_ready']
            del st.session_state['pdf_data']

# プレビュー表示（カラムの外＝全幅で表示）
if st.session_state.get('pdf_ready'):
    st.divider()
    st.subheader("レポートプレビュー")
    display_pdf(st.session_state['pdf_data'])

st.divider()
st.caption("© 2026 MOCAL株式会社 | GA4 Report Automation")
