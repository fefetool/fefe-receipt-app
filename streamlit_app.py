import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

# 頁面設定
st.set_page_config(page_title="收支憑證自動產生工具", layout="wide")
st.title("收支憑證自動產生工具")
st.markdown("請上傳 Excel 收支明細與 Word 憑證模板，點擊「開始產出憑證」。")

# 文件上傳區
col1, col2 = st.columns(2)

with col1:
    st.header("上傳 Excel 收支明細")
    uploaded_excel = st.file_uploader(
        "Drag and drop file here",
        type=["xlsx"],
        key="excel_uploader",
        label_visibility="collapsed"
    )
    st.caption("Limit 200MB per file • XLSX")

with col2:
    st.header("上傳 Word 憑證模板")
    uploaded_template = st.file_uploader(
        "Drag and drop file here",
        type=["docx"],
        key="word_uploader",
        label_visibility="collapsed"
    )
    st.caption("Limit 200MB per file • DOCX")

# 開始轉換按鈕
st.divider()
start_conversion = st.button("開始轉換並產出憑證", type="primary")

def convert_roc_date(roc_date):
    """將民國年日期轉換為datetime對象"""
    if pd.isna(roc_date):
        return None
        
    date_str = str(roc_date)
    
    # 處理多種日期格式
    match = re.search(r'(\d+)[-/年](\d+)[-/月](\d+)', date_str.replace(" ", ""))
    if match:
        roc_year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))
        
        # 民國年轉西元年
        if 100 < roc_year < 200:  # 合理民國年範圍
            western_year = roc_year + 1911
        else:
            western_year = roc_year
            
        try:
            return datetime(western_year, month, day)
        except:
            return None
    return None

def create_voucher_page(doc, record, is_income=True):
    """創建單一憑證頁面"""
    # 添加標題
    title = "收 入 憑 證 用 紙" if is_income else "支 出 憑 證 用 紙"
    doc.add_paragraph("台 日 產 業 技 術 合 作 促 進 會", style="Heading 1")
    heading = doc.add_paragraph(title)
    heading.style = doc.styles["Heading 2"]
    
    # 添加日期
    date_obj = record["日期"]
    roc_year = date_obj.year - 1911
    formatted_date = f"{roc_year}年{date_obj.month}月{date_obj.day}日"
    doc.add_paragraph(formatted_date)
    
    # 創建主表格
    table = doc.add_table(rows=2, cols=4)
    table.style = "Table Grid"
    
    # 設置表格標題行
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "憑 證 編 號"
    hdr_cells[1].text = "會 計 科 目"
    hdr_cells[2].text = "金    額"
    hdr_cells[3].text = "摘    要"
    
    # 填入數據
    row_cells = table.rows[1].cells
    row_cells[0].text = str(record["憑證編號"]) if pd.notna(record["憑證編號"]) else ""
    row_cells[1].text = str(record["科目"]) if pd.notna(record["科目"]) else ""
    amount = record["收入"] if is_income else record["支出"]
    row_cells[2].text = f"{int(amount):,}" if pd.notna(amount) else ""
    row_cells[3].text = str(record["摘要"]) if pd.notna(record["摘要"]) else ""
    
    # 添加簽名欄表格
    doc.add_paragraph()  # 空行
    sign_table = doc.add_table(rows=1, cols=4)
    sign_table.style = "Table Grid"
    sign_cells = sign_table.rows[0].cells
    sign_cells[0].text = "理 事 長"
    sign_cells[1].text = "秘 書 長"
    sign_cells[2].text = "副 秘 書 長"
    sign_cells[3].text = "製    單"
    
    # 添加說明文字
    doc.add_paragraph("說明：本單一式二聯，單位：新臺幣元。附單據。")
    
    # 設置全文字體
    set_document_font(doc)

def set_document_font(doc):
    """設置整份文件字體為標楷體"""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = "標楷體"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "標楷體"
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

if start_conversion:
    if not uploaded_excel or not uploaded_template:
        st.warning("請先上傳 Excel 收支明細和 Word 模板文件")
        st.stop()
    
    try:
        # 讀取 Excel 文件
        df = pd.read_excel(uploaded_excel, header=None)
        
        # 提取銀行交易記錄 (從第6行開始)
        bank_df = df.iloc[6:28, :8].copy()
        bank_df.columns = ["日期", "憑證編號", "科目", "摘要", "支出", "收入", "餘額", "備註"]
        bank_df["日期"] = bank_df["日期"].apply(convert_roc_date)
        bank_df = bank_df.dropna(subset=["日期"])
        
        # 提取現金交易記錄 (從第32行開始)
        cash_df = df.iloc[32:60, :7].copy()
        cash_df.columns = ["日期", "憑證編號", "科目", "摘要", "支出", "收入", "備註"]
        cash_df["日期"] = cash_df["日期"].apply(convert_roc_date)
        cash_df = cash_df.dropna(subset=["日期"])
        
    except Exception as e:
        st.error(f"讀取 Excel 文件時發生錯誤：{str(e)}")
        st.stop()

    try:
        # 創建新文件
        output_doc = Document()
        
        # 處理收入憑證
        income_records = bank_df[bank_df["收入"].notna() & (bank_df["收入"] != 0)]
        for idx, record in income_records.iterrows():
            create_voucher_page(output_doc, record, is_income=True)
            output_doc.add_page_break()
        
        # 處理支出憑證
        expense_records = pd.concat([
            bank_df[bank_df["支出"].notna() & (bank_df["支出"] != 0)],
            cash_df[cash_df["支出"].notna() & (cash_df["支出"] != 0)]
        ])
        for idx, record in expense_records.iterrows():
            create_voucher_page(output_doc, record, is_income=False)
            if idx < len(expense_records) - 1:  # 最後一筆不加分頁
                output_doc.add_page_break()
        
        # 保存結果
        output_buffer = BytesIO()
        output_doc.save(output_buffer)
        output_buffer.seek(0)
        
        st.success("憑證生成完成！共產生 {} 筆收入憑證和 {} 筆支出憑證。".format(
            len(income_records), len(expense_records)))
        st.download_button(
            label="下載憑證文件",
            data=output_buffer,
            file_name="收支憑證.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"處理 Word 模板時發生錯誤：{str(e)}")
        st.stop()
