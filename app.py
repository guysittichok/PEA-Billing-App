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
        result["L"] = float(energy.group(2).replace(",", ""))

    result["P"] = float(re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I) else ""


    return result

# Streamlit UI
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    data = [extract_exact_pea_bill(f) for f in uploaded_files]
    df = pd.DataFrame(data)
    st.data_editor(df, use_container_width=True)

st.markdown("---")
st.subheader("📥 ส่งออกไปยังไฟล์ Excel ต้นแบบ")
template_file = st.file_uploader("อัปโหลดไฟล์ Excel ต้นแบบ (.xlsx)", type=["xlsx"])

if template_file and not df.empty:
    if st.button("สร้างไฟล์ Excel พร้อมข้อมูล"):
        # 1. โหลดไฟล์ Excel ต้นแบบ
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active  # เลือก Sheet ที่ใช้งานอยู่ (ถ้ามีหลาย Sheet ให้ระบุชื่อ เช่น wb['Sheet1'])
        
        # 2. นำข้อมูลจาก df (ที่สกัดได้จาก PDF) ไปวางใน Excel
        # ตัวอย่าง: วางข้อมูลลงในช่อง E2, F2, ... (ปรับตำแหน่งตามไฟล์ต้นแบบของคุณ)
        for i, row in df.iterrows():
            # i+2 คือเริ่มที่แถวที่ 2 (สมมติแถว 1 เป็นหัวตาราง)
            ws[f'E{i+2}'] = row['E'] 
            ws[f'F{i+2}'] = row['F']
            # เพิ่มช่องอื่นๆ ได้ตามต้องการ เช่น ws[f'I{i+2}'] = row['I']
            
        # 3. เซฟไฟล์ลงหน่วยความจำ
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # 4. ปุ่มดาวน์โหลด
        st.download_button(
            label="ดาวน์โหลดไฟล์ Excel ผลลัพธ์",
            data=output,
            file_name="PEA_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
