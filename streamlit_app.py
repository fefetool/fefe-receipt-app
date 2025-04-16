import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re

# 頁面設定
st.set_page_config(page_title="收支憑證自動產生工具", layout="wide")
st.title("收支憑證自動產生工具")

# 函式：從日期字串提取民國年、月、日
def extract_date_parts(date_str):
    match = re.search(r'(\d+)年\s*(\d+)月\s*(\d+)日', date_str)
    if match:
        return match.group(1), match.group(2), match.group(3)
    return None, None, None

# 文件上傳區
st.header("請上傳 Excel 收支明細與 Word 樣板")

col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("上傳 Excel 收支明細", type=["xlsx"], key="excel_uploader")

with col2:
    uploaded_template = st.file_uploader("上傳 Word 憑證樣板", type=["docx"], key="word_uploader")

# 開始轉換按鈕
start_conversion = st.button("開始轉換並產出憑證")

if start_conversion:
    # 檢查文件是否已上傳
    if uploaded_excel is None or uploaded_template is None:
        st.warning("▲ 請上傳 Excel 收支明細與 Word 樣板")
        st.stop()

    try:
        # 讀取 Excel 文件
        df_raw = pd.read_excel(uploaded_excel, header=None)
        
        # 檢查是否有足夠的行數
        if len(df_raw) < 2:
            st.error("X Excel 文件格式錯誤：至少需要包含標題行和數據行")
            st.stop()
        
        # 提取日期和設置列名
        日期欄標題 = str(df_raw.iloc[0, 0])
        roc_year, month, day = extract_date_parts(日期欄標題)
        
        if None in (roc_year, month, day):
            st.error("X Excel 日期格式錯誤，請使用 '民國XXX年XX月XX日' 格式")
            st.stop()
            
        # 創建新的 DataFrame 並設置列名
        df = df_raw.iloc[2:].copy()
        df.columns = df_raw.iloc[1].tolist()
        
    except Exception as e:
        st.error(f"X 讀取 Excel 文件時發生錯誤：{e}")
        st.stop()

    try:
        # 讀取 Word 模板
        template_data = uploaded_template.read()
        st.session_state["template_data"] = template_data
        
        # 創建替換字典 (根據您的實際需求修改)
        replacements = {
            "YEAR": roc_year,
            "MONTH": month,
            "DAY": day,
            # 添加其他需要替換的欄位
        }
        
        # 處理 Word 文件
        output_doc = Document(BytesIO(template_data))
        
        # 替換模板中的標記
        for table in output_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, val in replacements.items():
                        if f"{{{{{key}}}}}" in cell.text:  # 使用 {{KEY}} 作為標記
                            cell.text = cell.text.replace(f"{{{{{key}}}}}", str(val))
        
        # 保存結果到記憶體
        output_buffer = BytesIO()
        output_doc.save(output_buffer)
        output_buffer.seek(0)
        
        st.success("● 憑證生成完成！")
        
        # 提供下載按鈕
        st.download_button(
            label="下載憑證",
            data=output_buffer,
            file_name="收支憑證.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"X 處理 Word 樣板時發生錯誤：{e}")
        st.stop()
