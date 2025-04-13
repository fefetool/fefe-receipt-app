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

欄位對應表 = {
    '日期': ['日期', '交易日期', '憑證日期', '入帳日'],
    '收入': ['收入', '收入金額', '收款金額'],
    '支出2': ['支出', '支出金額', '付款金額'],
    '用途': ['用途', '摘要', '說明', '用途說明'],
    '項目': ['項目', '科目', '分類', '費用類別']
}

if start_conversion:
    if uploaded_excel is None or uploaded_template is None:
        st.warning("⚠️ 請上傳 Excel 與 Word 樣板。")
        st.stop()

    for i in range(10):
        df_try = pd.read_excel(uploaded_excel, header=i)
        if any("日" in str(col) for col in df_try.columns):
            df_raw = df_try
            break
    else:
        st.error("❌ 無法找到有效的標題列，請確認檔案格式")
        st.stop()

    實際欄位 = {}
    for 標準欄, 可接受名 in 欄位對應表.items():
        for col in df_raw.columns:
            if any(name in str(col) for name in 可接受名):
                實際欄位[標準欄] = col
                break

    if not any(k in 實際欄位 for k in ['收入', '支出2']) or not all(k in 實際欄位 for k in ['日期', '用途', '項目']):
        st.error("❌ Excel 檔案欄位缺少，請確認包含：日期、收入 或 支出、用途、項目")
        st.stop()

    try:
        template_doc = Document(BytesIO(uploaded_template.read()))
        template_table = template_doc.tables[0]
    except Exception as e:
        st.error(f"❌ 無法讀取 Word 樣板：{e}")
        st.stop()

    st.success("✅ 已讀取收支明細，開始處理...")
    output_doc = Document()
    records = []
    counter_map = {}

    for _, row in df_raw.iterrows():
        try:
            if 實際欄位.get('收入') and pd.notna(row[實際欄位['收入']]):
                金額 = int(float(row[實際欄位['收入']]))
                類型 = 'A'
                表頭 = "收 入　憑　證  用　紙"
            elif 實際欄位.get('支出2') and pd.notna(row[實際欄位['支出2']]):
                金額 = int(float(row[實際欄位['支出2']]))
                類型 = 'B'
                表頭 = "支 出　憑　證  用　紙"
            else:
                continue

            roc_year, month, day = extract_date_parts(row[實際欄位['日期']])
            if roc_year == 0:
                continue
            date_code = f"{roc_year:03}{month:02}{day:02}"
            key = (date_code, 類型)
            counter_map[key] = counter_map.get(key, 0) + 1
            憑證編號 = f"{date_code}{類型}{counter_map[key]:02}"

            records.append({
                "憑證編號": 憑證編號,
                "科目": row.get(實際欄位['項目'], ''),
                "金額": 金額,
                "摘要": row.get(實際欄位['用途'], ''),
                "表頭": 表頭,
                "年": roc_year,
                "月": month,
                "日": day
            })
        except:
            continue

    if not records:
        st.warning("⚠️ 沒有可處理的資料。")
        st.stop()

    for rec in records:
        output_doc.add_paragraph("台 日 產 業 技 術 合 作 促 進 會").runs[0].font.size = Pt(13)
        sub = output_doc.add_paragraph(rec["表頭"])
        for run in sub.runs:
            run.font.size = Pt(16)
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

        table = output_doc.add_table(rows=len(template_table.rows), cols=len(template_table.columns))
        table.style = template_table.style
        table.autofit = False

        for i in range(len(template_table.rows)):
            for j in range(len(template_table.columns)):
                cell = table.cell(i, j)
                cell.text = template_table.cell(i, j).text
                apply_font(cell)

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
                apply_font(cell)

        date_p = output_doc.add_paragraph(f"{rec['年']} 年 {rec['月']} 月 {rec['日']} 日")
        for run in date_p.runs:
            run.font.size = Pt(11)
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

        output_doc.add_paragraph("………………憑………………證……………粘………………貼………………線……………")
        note = output_doc.add_paragraph("說明；本單一式一聯，單位：新臺幣元。附單據。")
        for run in note.runs:
            run.font.size = Pt(9)
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

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
