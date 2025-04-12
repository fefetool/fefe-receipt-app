import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from io import BytesIO
import datetime
import os

st.set_page_config(page_title="收支憑證自動產生工具", layout="wide")
st.title("\U0001F4C4 收支憑證自動產生工具")

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

st.markdown(" 請上傳 Excel 表單與 Word 樣板後，點擊『開始產出應證』。")

# 三個按鈕：左 - Excel，右 - Word，下 - 開始轉換
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("📂 上傳 Excel 收支明細", type=["xlsx"], key="excel")
with col2:
    uploaded_template = st.file_uploader("📄 上傳 Word 應證樣板", type=["docx"], key="word")

# 火箭啟動按鈕
start_conversion = st.button("🚀 開始轉換並產出應證")

# Function to apply font
def apply_font(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = '標潔體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標潔體')
            run.font.size = Pt(12)

def extract_date_parts(date_str):
    year, month, day = map(int, str(date_str).split('/'))
    return year, month, day

欄位對應表 = {
    '日期': ['日期', '交易日期'],
    '收入': ['收入', '收入金額'],
    '支出2': ['支出', '支出金額'],
    '用途': ['用途', '摘要', '說明'],
    '項目': ['項目', '科目', '分類']
}

if start_conversion:
    if uploaded_excel is None:
        st.warning("⚠️ 請先上傳 Excel 檔案")
        st.stop()
    df_raw = pd.read_excel(uploaded_excel, header=5)

    實際欄位 = {}
    for 標準欄, 可接受名 in 欄位對應表.items():
        for col in df_raw.columns:
            if any(name in str(col) for name in 可接受名):
                實際欄位[標準欄] = col
                break

    必要欄 = ['日期', '用途', '項目']
    if not any(k in 實際欄位 for k in ['收入', '支出2']) or not all(k in 實際欄位 for k in 必要欄):
        st.error("❌ Excel 檔案欄位缺少，請確認包含：日期、收入 或 支出、用途、項目")
        st.stop()

    if uploaded_template is None:
        st.error("❌ 請上傳 Word 應證樣板（.docx 檔案）")
        st.stop()

    try:
        template_data = uploaded_template.read()
        template_doc = Document(BytesIO(template_data))
    except Exception as e:
        st.error(f"❌ 無法讀取 Word 樣板：{e}")
        st.stop()

    st.success("✅ 已讀取收支明細，開始處理...")
    records = []
    counter_map = {}

    for _, row in df_raw.iterrows():
        if 實際欄位.get('收入') and pd.notna(row.get(實際欄位['收入'])):
            金額 = int(row[實際欄位['收入']])
            表頭 = "收 入　應　證  用　紙"
            類型 = 'A'
        elif 實際欄位.get('支出2') and pd.notna(row.get(實際欄位['支出2'])):
            金額 = int(row[實際欄位['支出2']])
            表頭 = "支 出　應　證  用　紙"
            類型 = 'B'
        else:
            continue

        try:
            roc_year, month, day = extract_date_parts(row[實際欄位['日期']])
        except:
            continue
        date_code = f"{roc_year:03}{month:02}{day:02}"
        key = (date_code, 類型)
        counter_map[key] = counter_map.get(key, 0) + 1
        seq = f"{counter_map[key]:02}"
        應證編號 = f"{date_code}{類型}{seq}"

        records.append({
            "應證編號": 應證編號,
            "科目": row.get(實際欄位['項目'], ''),
            "金額": 金額,
            "摘要": row.get(實際欄位['用途'], ''),
            "表頭": 表頭,
            "年": roc_year,
            "月": month,
            "日": day
        })

    if not records:
        st.warning("⚠️ 沒有可處理的資料。")
    else:
        try:
            template_table = template_doc.tables[0]
            output_doc = Document()

            for rec in records:
                table = output_doc.add_table(rows=len(template_table.rows), cols=len(template_table.columns))
                table.autofit = False

                for i in range(len(template_table.rows)):
                    for j in range(len(template_table.columns)):
                        cell = table.cell(i, j)
                        cell.text = template_table.cell(i, j).text
                        apply_font(cell)

                table.cell(0, 0).text = rec["表頭"]
                table.cell(2, 0).text = rec["應證編號"]
                table.cell(2, 1).text = rec["科目"]
                table.cell(2, 2).text = f"{rec['金額']:,}"
                table.cell(2, 3).text = rec["摘要"]

                for col in [0, 1, 2, 3]:
                    apply_font(table.cell(2, col))

                p = output_doc.add_paragraph(f"{rec['年']} 年 {rec['月']} 月 {rec['日']} 日")
                for run in p.runs:
                    run.font.name = '標潔體'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標潔體')
                    run.font.size = Pt(12)

                output_doc.add_paragraph()
                output_doc.add_page_break()

            buffer = BytesIO()
            output_doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="📅 下載產出應證 Word 檔",
                data=buffer,
                file_name="收支應證產出結果.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            with st.expander("📜 查看原始紀錄資料。"):
                st.dataframe(pd.DataFrame(records))
        except Exception as e:
            st.error(f"❌ 檔案產出錯誤：{e}")
