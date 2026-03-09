import streamlit as st
import os
import tempfile
import base64
from report_logic import generate_report
from report_logic_advanced import generate_report as generate_report_advanced

# ─────────────────────────────────────────────────────
# ユーティリティ
# ─────────────────────────────────────────────────────
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
st.title("📊 GA4 月次レポート自動生成ツール")
st.write("GA4からエクスポートしたCSVをアップロードするだけで、整形されたPDFレポートを作成します。")

st.divider()

# サイドバー設定
st.sidebar.header("📝 レポート設定")
company_name = st.sidebar.text_input("会社名", value="MOCAL株式会社")
staff_name = st.sidebar.text_input("担当者名", value="標 譲二")

st.sidebar.divider()
report_type = st.sidebar.radio(
    "レポートの種類",
    ["スタンダード (5ページ)", "詳細版 (10ページ)"],
    index=1,
    help="詳細版では、CVR解析やユーザー維持率、コンテンツカテゴリ分析が追加されます。"
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
        st.success("✅ CSVファイルが読み込めました！")
        
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
                        st.success("✨ レポートの生成が完了しました！")
                        
                        # ダウンロードボタン
                        st.download_button(
                            label="📥 PDFをダウンロードする",
                            data=pdf_data,
                            file_name=f"GA4レポート_{company_name}_{staff_name}.pdf",
                            mime="application/pdf"
                        )

                        # プレビュー表示（画面下部に大きく表示）
                        st.divider()
                        st.subheader("👁️ レポートプレビュー")
                        display_pdf(pdf_data)
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
    else:
        st.info("CSVファイルをアップロードしてください。")

st.divider()
st.caption("© 2026 MOCAL株式会社 | GA4 Report Automation")
