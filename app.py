import streamlit as st
import pdfplumber
import re
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="ระบบสกัดบิล PEA", layout="wide")

# --- นี่คือโค้ดประมวลผลที่คุณยืนยันว่าถูกต้อง ---
def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    result = {
        "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", 
        "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", 
        "O": "", "P": "", "Q": ""
    }

    # 1. Demand Charge
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if peak:
        result["C"] = float(peak.group(1).replace(",", "")); result["F"] = float(peak.group(2).replace(",", ""))
    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if pp:
        result["D"] = float(pp.group(1).replace(",", "")); result["G"] = float(pp.group(2).replace(",", ""))
    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # 2. Energy Usage
    for line in text.split('\n'):
        line_clean = line.strip()
        nums = re.findall(r"([\d,]+\.\d+)", line_clean)
        if not nums: continue
        if "พลังงานไฟฟ้า" in line_clean and "P" in line_clean and "PP" not in line_clean:
            result["I"] = float(nums[-1].replace(",", ""))
        elif "PP" in line_clean:
            result["J"] = float(nums[-1].replace(",", ""))
        elif "OP" in line_clean:
            result["K"] = float(nums[-1].replace(",", ""))

    # 3. Energy Cost & Others
    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", "")); result["L"] = float(energy.group(2).replace(",", ""))

    def get_val(p): 
        m = re.search(p, text, re.I)
        return float(m.group(1).replace(",", "")) if m else ""

    result["M"] = get_val(r'ค่า\s*Ft.*?([\d,]+\.\d+)')
    result["N"] = get_val(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)')
    result["P"] = get_val(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)')
    result["Q"] = get_val(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)')

    return result

# --- ส่วน UI และการกรอกลง Excel ---
st.title("⚡ ระบบสกัดข้อมูลและกรอก Excel")
uploaded_files = st.file_uploader("อัปโหลด PDF", accept_multiple_files=True)
template_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

if uploaded_files and template_file:
    if st.button("ประมวลผลและกรอกข้อมูล"):
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active # แก้ชื่อชีทถ้าต้องการ เช่น wb['Sheet1']
        
        for idx, f in enumerate(uploaded_files):
            data = extract_exact_pea_bill(f)
            # ปรับตำแหน่ง row ตามต้องการ (ตัวอย่างเริ่มที่แถว 5)
            row = 5 + idx 
            ws[f'C{row}'] = data['C']
            ws[f'D{row}'] = data['D']
            ws[f'E{row}'] = data['E']
            ws[f'I{row}'] = data['I']
            ws[f'J{row}'] = data['J']
            ws[f'K{row}'] = data['K']
            # เพิ่ม ws[...] = data['ตัวอักษรอื่น'] ได้ตามใจชอบ

        output = BytesIO()
        wb.save(output)
        st.download_button("🟢 ดาวน์โหลดไฟล์ Excel ที่กรอกข้อมูลแล้ว", data=output.getvalue(), file_name="Result.xlsx")
