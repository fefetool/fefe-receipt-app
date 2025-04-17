# streamlit_app.py - 主應用程式
import streamlit as st
import pandas as pd
from utils.template_parser import analyze_template, guess_excel_field
from utils.docx_generator import generate_voucher
from io import BytesIO
import hashlib

st.set_page_config(page_title="智慧憑證生成系統", page_icon="📄", layout="wide")
st.title("📄 台日產業技術合作促進會 - 收支憑證生成系統")

if 'template_mapping' not in st.session_state:
    st.session_state.template_mapping = {}
if 'saved_templates' not in st.session_state:
    st.session_state.saved_templates = {}

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. 上傳Excel收支明細")
    excel_file = st.file_uploader("選擇Excel檔案", type=["xlsx"], key="excel_uploader")

with col2:
    st.subheader("2. 上傳Word憑證模板")
    template_file = st.file_uploader("選擇Word模板", type=["docx"], key="template_uploader")

if excel_file and template_file:
    df = pd.read_excel(excel_file)
    template_hash = hashlib.md5(template_file.getvalue()).hexdigest()
    template_info = analyze_template(template_file)
    template_fields = template_info['available_fields']

    st.subheader("🔍 自動欄位分析")
    st.json(template_info)

    st.subheader("🧩 進階欄位對應設定")
    custom_mapping = {}
    for i in range(5):
        with st.expander(f"對應設定 #{i+1}", expanded=False):
            template_name = st.text_input(f"Word 欄位名稱 (例：會計科目)", key=f"t{i}")
            excel_name = st.selectbox(f"對應 Excel 欄位", options=df.columns, key=f"e{i}")
            if template_name and excel_name:
                custom_mapping[template_name] = excel_name

    field_mapping = {}
    for field in template_fields:
        if field in custom_mapping:
            mapped = custom_mapping[field]
        else:
            mapped = guess_excel_field(field, df.columns)
        if mapped:
            field_mapping[field] = {'template_field': field, 'excel_field': mapped}

    if st.button("🔄 生成憑證文件"):
        try:
            output = generate_voucher(df, template_file.getvalue(), field_mapping)
            st.success("✅ 憑證生成成功！")
            st.download_button("⬇️ 下載憑證", output, file_name="generated_receipts.docx")
        except Exception as e:
            st.error(f"❌ 發生錯誤: {str(e)}")
