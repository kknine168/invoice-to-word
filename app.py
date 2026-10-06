import streamlit as st
from docx import Document
from io import BytesIO

st.set_page_config(page_title="發票轉動支單系統", page_icon="📄", layout="centered")

st.title("📄 發票轉動支單系統 (免 API 版)")
st.write("請輸入動支用途與支出明細，系統將直接為您產出標準的 Word 動支單！")

# 1. 輸入介面
purpose = st.text_input("動支用途（例如：辦理資優班校外教學活動）")
uploaded_file = st.file_uploader("上傳發票照片或 PDF (備查用)", type=["jpg", "jpeg", "png", "pdf"])

st.subheader("請填寫支出明細")
item_name = st.text_input("品項名稱", placeholder="例如：教學實驗材料費")
quantity = st.number_input("數量", min_value=1, value=1)
price = st.number_input("單價 (元)", min_value=0, value=100)

total_amount = quantity * price
st.write(f"**計算總金額：NT$ {total_amount}**")

def generate_word_doc(purpose, item_name, quantity, price, total_amount):
    """產生標準 Word 動支單"""
    doc = Document()
    
    # 文件標題
    doc.add_heading('學校動支單 / 請示單', 0)
    
    # 填入用途
    doc.add_paragraph(f"動支用途：{purpose}")
    doc.add_paragraph("--------------------------------------------------")
    
    # 建立明細表格
    doc.add_heading('支出明細', level=2)
    table = doc.add_table(rows=1, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '品項名稱'
    hdr_cells[1].text = '數量'
    hdr_cells[2].text = '單價'
    hdr_cells[3].text = '小計'
    
    row_cells = table.add_row().cells
    row_cells[0].text = str(item_name)
    row_cells[1].text = str(quantity)
    row_cells[2].text = str(price)
    row_cells[3].text = str(total_amount)
        
    doc.add_paragraph(f"\n總計金額：NT$ {total_amount}")
    
    # 儲存至記憶體
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# 2. 產出按鈕
if st.button("產出 Word 動支單"):
    if not purpose:
        st.error("請輸入動支用途！")
    elif not item_name:
        st.error("請輸入品項名稱！")
    else:
        word_file = generate_word_doc(purpose, item_name, quantity, price, total_amount)
        st.success("動支單產生成功！")
        
        st.download_button(
            label="📥 下載 Word 動支單",
            data=word_file,
            file_name="動支單.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
