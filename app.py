import os
from io import BytesIO
import streamlit as st
from PIL import Image
import pdf2image
from openai import OpenAI
from docx import Document
import json
import base64

# 初始化 OpenAI 客戶端
client = OpenAI()

st.set_page_config(page_title="發票自動轉動支單系統", page_icon="📄", layout="centered")

st.title("📄 發票自動轉動支單系統")
st.write("請上傳發票（圖片或PDF檔）、輸入用途，並可選填上傳您的 Word 範本檔，系統將自動辨識並產出對應的 Word 動支單。")

# 1. 使用者輸入介面
purpose = st.text_input("請輸入動支用途（例如：辦理資優班校外教學活動）", placeholder="例如：購買資優班教學實驗器材")

uploaded_invoice = st.file_uploader("1. 上傳發票 (支援 JPG, PNG, PDF)", type=["jpg", "jpeg", "png", "pdf"])
uploaded_template = st.file_uploader("2. 上傳學校 Word 範本 (選填，若不上傳將使用預設格式)", type=["docx"])

def extract_invoice_data(image):
    """使用 OpenAI Vision API 解析發票內容"""
    buffered = BytesIO()
    image.save(buffered, format="JPEG")
    base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "system",
                "content": "你是一個專業的會計助理與發票辨識專家。請從發票中擷取品項、數量、單價與總金額。品項名稱請適度簡化。請以嚴格的 JSON 格式回傳，格式如下：\n"
                           '{\n  "items": [\n    {"name": "品項名稱", "quantity": 1, "price": 100, "total": 100}\n  ],\n  "total_amount": 100\n}'
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "請解析這張發票的品項、數量與金額。"},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                ]
            }
        ],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content)

def generate_word_doc(data, purpose, template_file=None):
    """根據使用者上傳的 Word 範本或預設格式產出文件"""
    if template_file:
        doc = Document(template_file)
    else:
        doc = Document()
        doc.add_heading('動支單 / 請示單', 0)
        doc.add_paragraph(f"動支用途：{purpose}")
        doc.add_heading('支出明細表', level=2)
        table = doc.add_table(rows=1, cols=4)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '品項名稱'
        hdr_cells[1].text = '數量'
        hdr_cells[2].text = '單價'
        hdr_cells[3].text = '小計'
        
        for item in data.get("items", []):
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get("name", ""))
            row_cells[1].text = str(item.get("quantity", ""))
            row_cells[2].text = str(item.get("price", ""))
            row_cells[3].text = str(item.get("total", ""))
            
        doc.add_paragraph(f"\n總計金額：NT$ {data.get('total_amount', 0)}")
    
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# 2. 執行按鈕
if st.button("開始辨識並產出 Word 檔"):
    if not uploaded_invoice:
        st.error("請先上傳發票檔案！")
    elif not purpose:
        st.error("請輸入動支用途！")
    else:
        with st.spinner("AI 正在辨識發票內容中，請稍候..."):
            try:
                if uploaded_invoice.type == "application/pdf":
                    images = pdf2image.convert_from_bytes(uploaded_invoice.read())
                    image = images[0]
                else:
                    image = Image.open(uploaded_invoice)
                
                invoice_data = extract_invoice_data(image)
                
                # 傳入範本（如果有的話）
                word_file = generate_word_doc(invoice_data, purpose, uploaded_template)
                
                st.success("辨識成功！請點擊下方按鈕下載 Word 檔。")
                
                st.download_button(
                    label="📥 下載產出的動支單 Word 檔",
                    data=word_file,
                    file_name="動支單_自動產出.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.write("### 辨識明細預覽：")
                st.json(invoice_data)
                
            except Exception as e:
                st.error(f"發生錯誤：{e}")
