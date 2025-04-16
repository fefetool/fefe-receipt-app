import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re
from datetime import datetime

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
    match = re.search(r'(\d+)[-/](\d+)[-/](\d+)', date_str)
    if match:
        roc_year = int(match.group(1))
        month = int(match.group(2))
        day = int(match.group(3))
        
        # 簡單判斷年份是否合理
        if roc_year > 150:  # 假設民國年>150表示可能是民國年
            western_year = roc_year + 1911
        else:
            western_year = roc_year
            
        try:
            return datetime(western_year, month, day)
        except:
            return None
    return None

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
        # 讀取 Word 模板
        template_data = uploaded_template.read()
        output_doc = Document(BytesIO(template_data))
        
        # 處理收入憑證
        income_records = bank_df[bank_df["收入"].notna() & (bank_df["收入"] != 0)]
        for idx, record in income_records.iterrows():
            if len(output_doc.tables) > 0:
                income_table = output_doc.tables[0]
                
                date_obj = record["日期"]
                roc_year = date_obj.year - 1911
                formatted_date = f"民國{roc_year}年{date_obj.month:02d}月{date_obj.day:02d}日"
                
                # 替換表格內容
                income_table.cell(0, 0).text = formatted_date  # 日期
                income_table.cell(2, 0).text = str(record["憑證編號"]) if pd.notna(record["憑證編號"]) else ""
                income_table.cell(2, 1).text = str(record["科目"]) if pd.notna(record["科目"]) else ""
                income_table.cell(2, 3).text = f"{int(record['收入']):,}"  # 金額
                income_table.cell(2, 5).text = str(record["摘要"]) if pd.notna(record["摘要"]) else ""
        
        # 處理支出憑證
        expense_records = pd.concat([
            bank_df[bank_df["支出"].notna() & (bank_df["支出"] != 0)],
            cash_df[cash_df["支出"].notna() & (cash_df["支出"] != 0)]
        ])
        for idx, record in expense_records.iterrows():
            if len(output_doc.tables) > 1:
                expense_table = output_doc.tables[1]
                
                date_obj = record["日期"]
                roc_year = date_obj.year - 1911
                formatted_date = f"民國{roc_year}年{date_obj.month:02d}月{date_obj.day:02d}日"
                
                # 替換表格內容
                expense_table.cell(0, 0).text = formatted_date  # 日期
                expense_table.cell(2, 0).text = str(record["憑證編號"]) if pd.notna(record.get("憑證編號")) else ""
                expense_table.cell(2, 1).text = str(record["科目"]) if pd.notna(record["科目"]) else ""
                expense_table.cell(2, 3).text = f"{int(record['支出']):,}"  # 金額
                expense_table.cell(2, 5).text = str(record["摘要"]) if pd.notna(record["摘要"]) else ""
        
        # 保存結果
        output_buffer = BytesIO()
        output_doc.save(output_buffer)
        output_buffer.seek(0)
        
        st.success("憑證生成完成！")
        st.download_button(
            label="下載憑證文件",
            data=output_buffer,
            file_name="收支憑證.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"處理 Word 模板時發生錯誤：{str(e)}")
        st.stop()
