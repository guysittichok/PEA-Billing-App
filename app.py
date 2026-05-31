import streamlit as st
import pdfplumber
import re
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="ระบบกรอกบิล PEA ลง Excel", layout="wide")
st.title("⚡ ระบบสกัดบิล PEA และกรอกลง Excel ต้นฉบับ")

# 1. ฟังก์ชันสกัดข้อมูล (ใช้ Logic เดิมของคุณ)
def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    result = {k: "" for k in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]}
    
    # [คง Logic เดิมของคุณไว้ทั้งหมด]
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if peak: result["C"] = float(peak.group(1).replace(",", "")); result["F"] = float(peak.group(2).replace(",", ""))
    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if pp: result["D"] = float(pp.group(1).replace(",", "")); result["G"] = float(pp.group(2).replace(",", ""))
    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
    if op: result["E"] = float(op.group(1).replace(",", ""))

    for line in text.split('\n'):
        line_clean = line.strip()
        nums = re.findall(r"([\d,]+\.\d+)", line_clean)
        if not nums: continue
        if "พลังงานไฟฟ้า" in line_clean and "P" in line_clean and "PP" not in line_clean: result["I"] = float(nums[-1].replace(",", ""))
        elif "PP" in line_clean: result["J"] = float(nums[-1].replace(",", ""))
        elif "OP" in line_clean: result["K"] = float(nums[-1].replace(",", ""))

    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy: result["O"] = float(energy.group(1).replace(",", "")); result["L"] = float(energy.group(2).replace(",", ""))
    
    # ค่าอื่นๆ
    for key, pattern in [("M", r'ค่า\s*Ft.*?([\d,]+\.\d+)'), ("N", r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)'), 
                         ("P", r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)'), 
                         ("Q", r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)')]:
        m = re.search(pattern, text, re.I)
        if m: result[key] = float(m.group(1).replace(",", ""))
    
    return result

# 2. UI และระบบกรอกข้อมูล
uploaded_files = st.file_uploader("1. อัปโหลดบิล PDF", type=["pdf"], accept_multiple_files=True)
template_file = st.file_uploader("2. อัปโหลดไฟล์ Excel ต้นฉบับ (.xlsx)", type=["xlsx"])

if uploaded_files and template_file:
    if st.button("ประมวลผลและกรอกลง Excel"):
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active 
        
        for idx, f in enumerate(uploaded_files):
            data = extract_exact_pea_bill(f)
            # กำหนดตำแหน่งลงเซลล์ (แก้ตรงนี้ให้ตรงกับ Excel ของคุณ)
            row = 5 + idx # เริ่มกรอกแถวที่ 5
            ws[f'C{row}'] = data['C']; ws[f'D{row}'] = data['D']; ws[f'E{row}'] = data['E']
            ws[f'F{row}'] = data['F']; ws[f'G{row}'] = data['G']; ws[f'I{row}'] = data['I']
            ws[f'J{row}'] = data['J']; ws[f'K{row}'] = data['K']; ws[f'L{row}'] = data['L']
        
        # ดาวน์โหลดไฟล์
        output = BytesIO()
        wb.save(output)
        st.download_button("🟢 ดาวน์โหลดไฟล์ที่กรอกข้อมูลแล้ว", data=output.getvalue(), file_name="Report_Filled.xlsx")
