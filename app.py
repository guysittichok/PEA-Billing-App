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

# ส่วนนี้คือการกรอกข้อมูลลง Excel โดยคงสูตรเดิมไว้
if template_file and st.button("สร้างไฟล์ Excel พร้อมข้อมูล"):
    try:
        # โหลดไฟล์ Excel ต้นแบบ
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active 
        
        # วนลูปเช็คแถวใน Excel (เริ่มจากแถว 20 ตามภาพของคุณ)
        # วนลูปเช็คแถวใน Excel
        for row_idx in range(20, ws.max_row + 1):
            excel_key = str(ws[f'A{row_idx}'].value).strip()
            if excel_key == "None" or excel_key == "": continue
            
            match_row = df[df['ชื่อไฟล์'].apply(lambda x: excel_key in str(x))]
            
            if not match_row.empty:
                row = match_row.iloc[0]
                
                # ฟังก์ชันช่วยเขียนค่าแบบบังคับเป็นตัวเลข
                def write_number(ws, cell_pos, value):
                    try:
                        # แปลงเป็น float ถ้าเป็นตัวเลข/string ที่เป็นตัวเลข
                        val = float(str(value).replace(',', ''))
                        ws[cell_pos] = val
                        ws[cell_pos].number_format = '0.00' # ตั้งรูปแบบทศนิยม 2 ตำแหน่ง
                    except:
                        ws[cell_pos] = 0 # ถ้าแปลงไม่ได้ให้เป็น 0 เพื่อไม่ให้สูตร Error

                # เขียนข้อมูลทุกคอลัมน์โดยใช้ฟังก์ชันด้านบน
                write_number(ws, f'C{row_idx}', row.get('C', 0))
                write_number(ws, f'D{row_idx}', row.get('D', 0))
                write_number(ws, f'E{row_idx}', row.get('E', 0))
                write_number(ws, f'F{row_idx}', row.get('F', 0))
                write_number(ws, f'G{row_idx}', row.get('G', 0))
                write_number(ws, f'I{row_idx}', row.get('I', 0))
                write_number(ws, f'J{row_idx}', row.get('J', 0))
                write_number(ws, f'K{row_idx}', row.get('K', 0))
                write_number(ws, f'L{row_idx}', row.get('L', 0))
                write_number(ws, f'P{row_idx}', row.get('P', 0))
        
        # เตรียมไฟล์สำหรับดาวน์โหลด
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.success("กรอกข้อมูลลงไฟล์ต้นแบบเรียบร้อยแล้ว!")
        st.download_button(
            label="📥 ดาวน์โหลด Excel ที่กรอกข้อมูลเสร็จแล้ว",
            data=output,
            file_name="PEA_Result_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: {e}")
