import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("⚡ ระบบสกัดบิล PEA และกรอกลง Excel")

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

uploaded_files = st.file_uploader("1. อัปโหลดบิล PDF", accept_multiple_files=True)
template_file = st.file_uploader("2. อัปโหลดไฟล์ Excel (.xlsx)", type=["xlsx"])

if uploaded_files and template_file:
    if st.button("เริ่มประมวลผลและสร้างไฟล์"):
        # 1. สกัดข้อมูล
        all_data = [extract_exact_pea_bill(f) for f in uploaded_files]
        df_extracted = pd.DataFrame(all_data)
        
        # 2. อ่านไฟล์ Excel ต้นฉบับ
        df_template = pd.read_excel(template_file)
        
        # 3. กรอกข้อมูล: สมมติว่าต้องการเอาข้อมูลไป "ต่อท้าย" ใน Excel
        # ตรงนี้คือการนำข้อมูลที่สกัดได้ไป "แปะ" ลงใน DataFrame เดิม
        df_combined = pd.concat([df_template, df_extracted], ignore_index=True)
        
        # 4. สร้างไฟล์ใหม่
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_combined.to_excel(writer, index=False)
        
        st.success("ประมวลผลเสร็จแล้ว!")
        st.download_button("🟢 ดาวน์โหลด Excel ที่มีข้อมูลแล้ว", data=output.getvalue(), file_name="Completed_Report.xlsx")
