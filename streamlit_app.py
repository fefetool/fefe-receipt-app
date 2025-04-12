
# 注意：此程式需在安裝有 streamlit 的本地環境中執行。
# 安裝指令：pip install streamlit pandas python-docx openpyxl

try:
    import streamlit as st
    import pandas as pd
    from docx import Document
    from docx.shared import Pt
    from docx.oxml.ns import qn
    from io import BytesIO
    import datetime
except ModuleNotFoundError as e:
    import sys
    print("\n🚨 缺少必要套件：", e.name)
    print("請先安裝所需套件：pip install streamlit pandas python-docx openpyxl")
    sys.exit(1)

st.set_page_config(page_title="收支憑證自動產生工具", layout="centered")
st.title("📄 收支憑證自動產生工具")
st.markdown("請上傳格式正確的 Excel 收支明細，將自動生成會計版 Word 收支憑證。")

# Function to apply font
def apply_font(cell):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
            run.font.size = Pt(12)

# Function to extract date from voucher ID
def extract_date_parts(date_str):
    year, month, day = map(int, date_str.split('/'))
    return year, month, day

# Upload section
uploaded_file = st.file_uploader("請上傳收支明細 Excel 檔案：", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=3)  # 以第4列為標題列

    if '日期' not in df_raw.columns or ('收入' not in df_raw.columns and '支出2' not in df_raw.columns):
        st.error("❌ 欄位缺少，請確認包含：日期、收入 或 支出2、用途、項目")
    else:
        st.success("✅ 已讀取收支明細，開始處理...")

        records = []
        counter_map = {}  # key = (date, A/B) → counter

        for _, row in df_raw.iterrows():
            if pd.notna(row.get('收入')):
                金額 = int(row['收入'])
                表頭 = "收 入　憑　證  用　紙"
                類型 = 'A'
            elif pd.notna(row.get('支出2')):
                金額 = int(row['支出2'])
                表頭 = "支 出　憑　證  用　紙"
                類型 = 'B'
            else:
                continue

            # 編號產生
            try:
                roc_year, month, day = extract_date_parts(str(row['日期']))
            except:
                continue
            date_code = f"{roc_year:03}{month:02}{day:02}"
            key = (date_code, 類型)
            counter_map[key] = counter_map.get(key, 0) + 1
            seq = f"{counter_map[key]:02}"
            憑證編號 = f"{date_code}{類型}{seq}"

            records.append({
                "憑證編號": 憑證編號,
                "科目": row.get('項目', ''),
                "金額": 金額,
                "摘要": row.get('用途', ''),
                "表頭": 表頭,
                "年": roc_year,
                "月": month,
                "日": day
            })

        if not records:
            st.warning("⚠️ 找不到可用的收入或支出資料。")
        else:
            output_doc = Document()
            template_doc = Document("憑證樣板.docx")
            template_table = template_doc.tables[0]

            for rec in records:
                table = output_doc.add_table(rows=len(template_table.rows), cols=len(template_table.columns))
                table.autofit = False

                for i in range(len(template_table.rows)):
                    for j in range(len(template_table.columns)):
                        cell = table.cell(i, j)
                        cell.text = template_table.cell(i, j).text
                        apply_font(cell)

                table.cell(0, 0).text = rec["表頭"]
                table.cell(2, 0).text = rec["憑證編號"]
                table.cell(2, 1).text = rec["科目"]
                table.cell(2, 2).text = f"{rec['金額']:,}"
                table.cell(2, 3).text = rec["摘要"]

                for col in [0, 1, 2, 3]:
                    apply_font(table.cell(2, col))

                p = output_doc.add_paragraph(f"{rec['年']} 年 {rec['月']} 月 {rec['日']} 日")
                for run in p.runs:
                    run.font.name = '標楷體'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')
                    run.font.size = Pt(12)

                output_doc.add_paragraph()
                output_doc.add_page_break()

            buffer = BytesIO()
            output_doc.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="📥 下載收支憑證 Word 檔",
                data=buffer,
                file_name="114_3月收支憑證.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
