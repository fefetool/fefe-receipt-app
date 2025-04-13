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

def apply_font(cell, font_size=11):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(font_size)

def extract_date_parts(date_str):
    try:
        if isinstance(date_str, datetime.date):
            return date_str.year - 1911, date_str.month, date_str.day
        year, month, day = map(int, str(date_str).split('/'))
        return year, month, day
    except:
        return 0, 0, 0

if start_conversion:
    if uploaded_excel is None or uploaded_template is None:
        st.warning("⚠️ 請上傳 Excel 與 Word 樣板。")
        st.stop()

    df_raw = pd.read_excel(uploaded_excel, header=None)
    try:
        日期欄標題 = df_raw.iloc[0, 0]
        roc_year, month, day = extract_date_parts(日期欄標題)
        df_raw.columns = df_raw.iloc[1]  # 將第二列設為欄位名稱
        df_raw = df_raw[2:]  # 資料從第三列開始
    except:
        st.error("❌ Excel 日期欄與標題列格式不符，請依照標準範本製作。")
        st.stop()

    try:
        template_doc = Document(BytesIO(uploaded_template.read()))
        template_table = template_doc.tables[0]
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
        table = output_doc.add_table(rows=len(template_table.rows), cols=len(template_table.columns))
        table.style = template_table.style

        for i, tmpl_row in enumerate(template_table.rows):
            for j, tmpl_cell in enumerate(tmpl_row.cells):
                new_cell = table.cell(i, j)
                new_cell.text = tmpl_cell.text
                apply_font(new_cell)

        for row in table.rows:
            for cell in row.cells:
                text = cell.text
                if "憑證編號" in text:
                    cell.text = rec["憑證編號"]
                elif "會計科目" in text:
                    cell.text = rec["科目"]
                elif "金額" in text:
                    cell.text = f"{rec['金額']:,}"
                elif "摘要" in text:
                    cell.text = rec["摘要"]
                elif "日期" in text:
                    cell.text = f"{rec['年']} 年 {rec['月']} 月 {rec['日']} 日"
                apply_font(cell)

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
