import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO
import openpyxl

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบสกัดข้อมูลบิลค่าไฟฟ้า PEA")

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", 
        "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", 
        "O": "", "P": "", "Q": ""
    }

    # 1. Demand Charge (C, D, E, F, G)
    # ใช้ Regex ที่ระบุพิกัดช่องชัดเจน
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if peak:
        result["C"] = float(peak.group(1).replace(",", ""))
        result["F"] = float(peak.group(2).replace(",", ""))

    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if pp:
        result["D"] = float(pp.group(1).replace(",", ""))
        result["G"] = float(pp.group(2).replace(",", ""))

    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # 2. Energy Usage (I, J, K) - ดึงแบบระบุบรรทัด "พลังงานไฟฟ้า"
    # แยกบรรทัดชัดเจนเพื่อไม่ให้สลับช่องกัน
    for line in text.split('\n'):
        line_clean = line.strip()
        nums = re.findall(r"([\d,]+\.\d+)", line_clean)
        if not nums: continue
        
        # ค้นหา Peak: บรรทัดต้องมีคำว่า "พลังงานไฟฟ้า" และ "P"
        if "พลังงานไฟฟ้า" in line_clean and "P" in line_clean and "PP" not in line_clean:
            result["I"] = float(nums[-1].replace(",", ""))
            
        # ค้นหา PP: บรรทัดต้องมี "PP"
        elif "PP" in line_clean:
            result["J"] = float(nums[-1].replace(",", ""))
            
        # ค้นหา OP: บรรทัดต้องมี "OP"
        elif "OP" in line_clean:
            result["K"] = float(nums[-1].replace(",", ""))

    # 3. Energy Cost & Others
    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", ""))
        result["L"] = float(energy.group(2).replace(",", ""))

    result["M"] = float(re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I) else ""
    result["N"] = float(re.search(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text, re.I) else ""
    result["P"] = float(re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I) else ""
    result["Q"] = float(re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text, re.I) else ""

    return result

# Streamlit UI
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    data = [extract_exact_pea_bill(f) for f in uploaded_files]
    df = pd.DataFrame(data)
    st.data_editor(df, use_container_width=True)

st.write("---")
st.subheader("ส่วนการกรอกข้อมูลลง Excel")
template_file = st.file_uploader("อัปโหลดไฟล์ Excel ต้นฉบับ (.xlsx) เพื่อกรอกค่า", type=["xlsx"])

if template_file:
    if st.button("กรอกข้อมูลลง Excel"):
        # โหลดไฟล์ Excel ต้นฉบับ
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active # เลือกชีทที่ใช้งาน (แก้ชื่อชีทได้ เช่น wb['Sheet1'])
        
        # นำข้อมูลจาก df (ที่ได้จากโค้ดเดิมของคุณ) มาวนลูปกรอก
        # สมมติว่าในตาราง df ของคุณมีข้อมูลเรียงตามลำดับไฟล์ที่อัปโหลด
        for idx, row in df.iterrows():
            # กำหนดแถวที่จะกรอก (เช่น เริ่มที่แถว 5)
            row_idx = 5 + idx 
            
            # --- ตรงนี้คือจุดที่คุณต้องแก้พิกัดให้ตรงกับช่องใน Excel ของคุณ ---
            # ตัวอย่าง: ค่า C ของโปรแกรม ไปลงช่อง D ของ Excel
            ws[f'D{row_idx}'] = row['C']
            ws[f'E{row_idx}'] = row['D']
            ws[f'F{row_idx}'] = row['E']
            ws[f'I{row_idx}'] = row['I']
            ws[f'J{row_idx}'] = row['J']
            ws[f'K{row_idx}'] = row['K']
            ws[f'L{row_idx}'] = row['L']
            # คุณสามารถเพิ่ม ws[f'ช่อง{row_idx}'] = row['ตัวอักษร'] ได้ตามต้องการ
            
        # สร้างไฟล์ใหม่ให้ดาวน์โหลด
        output = BytesIO()
        wb.save(output)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel ที่กรอกข้อมูลแล้ว",
            data=output.getvalue(),
            file_name="PEA_Filled_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
