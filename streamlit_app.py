import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from io import BytesIO
import datetime

st.set_page_config(page_title="收支憑證自動產生工具", layout="wide")
st.title("📄 收支憑證自動產生工具")

# 密碼保護
if "authenticated" not in st.session_state:
    password = st.text_input("請輸入密碼以進入系統：", type="password")
    if password == "FEFE":
        st.session_state.authenticated = True
        st.success("✅ 密碼正確，請繼續操作。")
        st.rerun()
    elif password:
        st.error("❌ 密碼錯誤，請再試一次。")
    st.stop()

st.markdown("請上傳 Excel 表單與 Word 樣板後，點擊『開始產出憑證』。")

col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📂 上傳 Excel 收支明細", type=["xlsx"], key="excel")
with col2:
    uploaded_template = st.file_uploader("📄 上傳 Word 憑證樣板", type=["docx"], key="word")

start_conversion = st.button("🚀 開始轉換並產出憑證")

def extract_date_parts(date_str):
    try:
        if isinstance(date_str, datetime.date):
            return date_str.year - 1911, date_str.month, date_str.day
        year, month, day = map(int, str(date_str).split('/'))
        return year, month, day
    except:
        return 0, 0, 0

def replace_placeholder(paragraphs, placeholder, new_text):
    for paragraph in paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, new_text)

if start_conversion:
    if uploaded_excel is None or uploaded_template is None:
        st.warning("⚠️ 請上傳 Excel 與 Word 樣板。")
        st.stop()

    df_raw = pd.read_excel(uploaded_excel, header=None)
    try:
        日期欄標題 = df_raw.iloc[0, 0]
        roc_year, month, day = extract_date_parts(日期欄標題)
        df_raw.columns = df_raw.iloc[1]
        df_raw = df_raw[2:]
    except:
        st.error("❌ Excel 日期欄與標題列格式不符，請依照標準範本製作。")
        st.stop()

    try:
        template_data = uploaded_template.read()
        st.session_state["template_data"] = template_data
    except Exception as e:
        st.error(f"❌ 無法讀取 Word 憑證樣板：{e}")
        st.stop()

    st.success("✅ 已讀取收支明細，開始處理...")
    output_doc = Document()
    records = []

    for _, row in df_raw.iterrows():
        try:
            憑證編號 = str(row.get("憑證編號", "")).strip()
            科目 = str(row.get("會計科目", "")).strip()
            金額 = int(float(row.get("金額", 0)))
            摘要 = str(row.get("摘要", "")).strip()
        except:
            continue

        records.append({
            "憑證編號": 憑證編號,
            "科目": 科目,
            "金額": 金額,
            "摘要": 摘要,
            "年": roc_year,
            "月": month,
            "日": day
        })

    if not records:
        st.warning("⚠️ 沒有可處理的資料。")
        st.stop()

    for rec in records:
        template_doc = Document(BytesIO(st.session_state["template_data"]))

        replace_placeholder(template_doc.paragraphs, "{{憑證編號}}", rec["憑證編號"])
        replace_placeholder(template_doc.paragraphs, "{{會計科目}}", rec["科目"])
        replace_placeholder(template_doc.paragraphs, "{{金額}}", f"{rec['金額']:,}")
        replace_placeholder(template_doc.paragraphs, "{{摘要}}", rec["摘要"])
        replace_placeholder(template_doc.paragraphs, "{{日期}}", f"{rec['年']} 年 {rec['月']} 月 {rec['日']} 日")

        for element in template_doc.element.body:
            output_doc.element.body.append(element)

        output_doc.add_page_break()

    buffer = BytesIO()
    output_doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="📥 下載產出憑證 Word 檔",
        data=buffer,
        file_name="收支憑證產出結果.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    with st.expander("📋 查看原始紀錄資料"):
        st.dataframe(pd.DataFrame(records))
