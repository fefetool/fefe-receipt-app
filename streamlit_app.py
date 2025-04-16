import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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

def set_cell_border(cell, border_style="single", border_size=4, border_color="000000"):
    """設定表格儲存格邊框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    tcBorders = OxmlElement('w:tcBorders')
    
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_element = OxmlElement(f'w:{edge}')
        edge_element.set(qn('w:val'), border_style)
        edge_element.set(qn('w:sz'), str(border_size))
        edge_element.set(qn('w:color'), border_color)
        tcBorders.append(edge_element)
    
    tcPr.append(tcBorders)

def create_voucher_page(doc, record, is_income=True):
    """創建單一憑證頁面（完全符合樣本格式）"""
    # 1. 添加標題「台日產業技術合作促進會」
    title = doc.add_paragraph()
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = title.add_run("台  日  產  業  技  術  合  作  促  進  會")
    run.font.name = "標楷體"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0, 0, 0)  # 黑色
    
    # 2. 添加「收入憑證用紙」或「支出憑證用紙」
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    voucher_text = "收  入  憑  證  用  紙" if is_income else "支  出  憑  證  用  紙"
    run = subtitle.add_run(voucher_text)
    run.font.name = "標楷體"
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)  # 黑色
    
    # 3. 添加日期 (格式：114年3月5日)
    date_obj = record["日期"]
    roc_year = date_obj.year - 1911
    formatted_date = f"{roc_year}年{date_obj.month}月{date_obj.day}日"
    date_para = doc.add_paragraph(formatted_date)
    date_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in date_para.runs:
        run.font.name = "標楷體"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        run.font.size = Pt(12)
    
    # 4. 創建主表格
    table = doc.add_table(rows=2, cols=4)
    table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    table.style = "Table Grid"
    
    # 設定表格寬度 (17.5公分)
    table.width = Cm(17.5)
    
    # 設定欄寬
    table.columns[0].width = Cm(3.5)  # 憑證編號
    table.columns[1].width = Cm(3.5)  # 會計科目
    table.columns[2].width = Cm(3.5)  # 金額
    table.columns[3].width = Cm(7.0)  # 摘要
    
    # 設定表格邊框
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell)
    
    # 填入標題行 (第一列)
    hdr_cells = table.rows[0].cells
    headers = ["憑  證  編  號", "會  計  科  目", "金　　　　額", "摘　　　　要"]
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in hdr_cells[i].paragraphs[0].runs:
            run.font.name = "標楷體"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(12)
            run.font.bold = True
    
    # 填入數據行 (第二列)
    row_cells = table.rows[1].cells
    row_cells[0].text = str(record["憑證編號"]) if pd.notna(record["憑證編號"]) else ""
    row_cells[1].text = str(record["科目"]) if pd.notna(record["科目"]) else ""
    amount = record["收入"] if is_income else record["支出"]
    row_cells[2].text = f"{int(amount):,}" if pd.notna(amount) else ""
    row_cells[3].text = str(record["摘要"]) if pd.notna(record["摘要"]) else ""
    
    # 設定數據行格式
    row_cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 憑證編號置中
    row_cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 會計科目置中
    row_cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 金額置中
    row_cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT    # 摘要置左
    
    for cell in row_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = "標楷體"
                run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                run.font.size = Pt(12)
    
    # 5. 添加簽名欄位
    doc.add_paragraph()  # 空行
    
    sign_table = doc.add_table(rows=1, cols=4)
    sign_table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # 設定簽名表格寬度
    sign_table.width = Cm(17.5)
    
    # 設定簽名表格欄寬
    sign_table.columns[0].width = Cm(4.0)
    sign_table.columns[1].width = Cm(4.0)
    sign_table.columns[2].width = Cm(4.0)
    sign_table.columns[3].width = Cm(5.5)
    
    # 填入簽名欄位
    sign_cells = sign_table.rows[0].cells
    sign_texts = ["理  事  長", "秘  書  長", "副  秘  書  長", "製　　　單"]
    for i, text in enumerate(sign_texts):
        sign_cells[i].text = text
        sign_cells[i].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in sign_cells[i].paragraphs[0].runs:
            run.font.name = "標楷體"
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(12)
    
    # 6. 添加憑證粘貼線
    doc.add_paragraph()  # 空行
    line_para = doc.add_paragraph()
    line_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    line_run = line_para.add_run("..................憑..................證...............粘..................貼..................線...................")
    line_run.font.name = "標楷體"
    line_run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
    line_run.font.size = Pt(9)
    
    # 7. 添加說明文字
    note_para = doc.add_paragraph("說明：本單一式二聯，單位：新臺幣元。附單據。")
    note_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in note_para.runs:
        run.font.name = "標楷體"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        run.font.size = Pt(10)
    
    # 添加分頁符
    doc.add_page_break()

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
        
        # 設定文件預設字體
        style = output_doc.styles['Normal']
        font = style.font
        font.name = '標楷體'
        font._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
        
        # 處理收入憑證
        income_records = bank_df[bank_df["收入"].notna() & (bank_df["收入"] != 0)]
        for idx, record in income_records.iterrows():
            create_voucher_page(output_doc, record, is_income=True)
        
        # 處理支出憑證
        expense_records = pd.concat([
            bank_df[bank_df["支出"].notna() & (bank_df["支出"] != 0)],
            cash_df[cash_df["支出"].notna() & (cash_df["支出"] != 0)]
        ])
        for idx, record in expense_records.iterrows():
            create_voucher_page(output_doc, record, is_income=False)
        
        # 移除最後一頁多餘的分頁符
        if len(output_doc.paragraphs) > 0:
            last_paragraph = output_doc.paragraphs[-1]
            if last_paragraph.runs and last_paragraph.runs[-1].text == "":
                output_doc._body.remove(last_paragraph._element)
        
        # 保存結果
        output_buffer = BytesIO()
        output_doc.save(output_buffer)
        output_buffer.seek(0)
        
        st.success(f"憑證生成完成！共產生 {len(income_records)} 筆收入憑證和 {len(expense_records)} 筆支出憑證。")
        st.download_button(
            label="下載憑證文件",
            data=output_buffer,
            file_name="收支憑證.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"處理 Word 模板時發生錯誤：{str(e)}")
        st.stop()
