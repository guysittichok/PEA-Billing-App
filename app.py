import streamlit as st
import pdfplumber
import re
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบสกัดข้อมูลบิลค่าไฟฟ้า PEA และบันทึกลง Excel")

# ฟังก์ชันสกัดข้อมูล
def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        tables = page.extract_tables()

    result = {"C": 0.0, "D": 0.0, "E": 0.0, "F": 0.0, "G": 0.0, "H": 0.0, 
              "I": 0.0, "J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0, 
              "O": 0.0, "P": 0.0, "Q": 0.0}

    # 1. ดึง Demand Charge
    peak = re.search(r'Peak\s+([\d,]+\.\d+)', text, re.I)
    if peak: result["C"] = float(peak.group(1).replace(",", ""))
    
    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)', text, re.I)
    if pp: result["D"] = float(pp.group(1).replace(",", ""))

    # 2. ดึงข้อมูลจากตารางสำหรับ I, J, K
    for table in tables:
        for row in table:
            row_str = " ".join([str(cell) for cell in row if cell])
            if "พลังงานไฟฟ้า" in row_str:
                nums = re.findall(r"[\d,]+\.\d+", row_str)
                if len(nums) >= 3:
                    result["I"] = float(nums[0].replace(",", ""))
                    result["J"] = float(nums[1].replace(",", ""))
                    result["K"] = float(nums[2].replace(",", ""))
    
    # 3. ดึงค่าอื่นๆ
    q_match = re.search(r'รวมเงินค่าไฟฟ้า.*?([\d,]+\.\d+)', text, re.I)
    if q_match: result["Q"] = float(q_match.group(1).replace(",", ""))

    return result

# ฟังก์ชันแก้ไขไฟล์ Excel
def update_excel_file(template_path, extracted_data):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    # Map คอลัมน์ C(3) ถึง Q(17)
    column_map = {
        'C': 3, 'D': 4, 'E': 5, 'F': 6, 'G': 7, 'H': 8, 
        'I': 9, 'J': 10, 'K': 11, 'L': 12, 'M': 13, 'N': 14, 
        'O': 15, 'P': 16, 'Q': 17
    }

    for col_key, col_index in column_map.items():
        ws.cell(row=20, column=col_index).value = extracted_data[col_key]

    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- หน้า UI ---
uploaded_file = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"])

if uploaded_file:
    data = extract_exact_pea_bill(uploaded_file)
    st.write("### ข้อมูลที่สกัดได้:")
    st.json(data)

    if st.button("บันทึกลง template.xlsx และดาวน์โหลด"):
        try:
            excel_binary = update_excel_file("template.xlsx", data)
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Excel ที่อัปเดตแล้ว",
                data=excel_binary,
                file_name="ไฟฟ้า_อัปเดต_Row20.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("บันทึกข้อมูลเรียบร้อย!")
        except FileNotFoundError:
            st.error("ไม่พบไฟล์ 'template.xlsx' กรุณาตรวจสอบว่ามีไฟล์นี้อยู่ในโฟลเดอร์เดียวกันแล้ว")
