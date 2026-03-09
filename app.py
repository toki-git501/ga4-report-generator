import streamlit as st
import os
import tempfile
import base64
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
    """PDFをbase64エンコードしてiframeで表示する"""
    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="1000" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────
# Streamlit UI 設定
# ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="GA4 Report Generator",
    page_icon="📊",
    layout="wide"
)

# カスタムCSS（MOCALカラー）
st.markdown("""
    <style>
    .main {
        background-color: #F5F8FC;
    }
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
company_name = st.sidebar.text_input("会社名", value="MOCAL株式会社")
staff_name = st.sidebar.text_input("担当者名", value="標 譲二")

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

        if st.button("🚀 PDFレポートを生成する"):
            with st.spinner("レポートを生成中です。しばらくお待ちください..."):
                try:
                    # 一時ディレクトリにファイルを保存
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        # CSV保存
                        csv_path = os.path.join(tmp_dir, "input.csv")
                        with open(csv_path, "wb") as f:
                            f.write(uploaded_csv.getbuffer())

                        # ロゴ保存（あれば）
                        logo_path = None
                        if uploaded_logo:
                            logo_path = os.path.join(tmp_dir, "logo.png")
                            with open(logo_path, "wb") as f:
                                f.write(uploaded_logo.getbuffer())

                        # 出力パス
                        output_pdf_path = os.path.join(tmp_dir, "report.pdf")

                        # レポート生成実行
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

                        # 生成されたPDFを読み込む
                        with open(output_pdf_path, "rb") as pdf_file:
                            pdf_data = pdf_file.read()

                        st.balloons()
                        st.success("レポートの生成が完了しました！")

                        # ダウンロードボタン
                        st.download_button(
                            label="PDFをダウンロードする",
                            data=pdf_data,
                            file_name=f"GA4レポート_{company_name}_{staff_name}.pdf",
                            mime="application/pdf"
                        )

                        # プレビュー表示（画面下部に大きく表示）
                        st.divider()
                        st.subheader("レポートプレビュー")
                        display_pdf(pdf_data)
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
    else:
        st.info("CSVファイルをアップロードしてください。")

st.divider()
st.caption("© 2026 MOCAL株式会社 | GA4 Report Automation")
