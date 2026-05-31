import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("⚡ ระบบสกัดข้อมูลบิล PEA (โครงสร้างเสถียร)")

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", 
        "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", 
        "O": "", "P": "", "Q": ""
    }

    # 1. กลุ่ม Demand (C, D, E, F, G)
    for line in text.split('\n'):
        # ค้นหา Peak/Partial/Off แล้วดึงตัวเลขในบรรทัดนั้น
        nums = re.findall(r"[\d,]+\.\d+", line)
        if not nums: continue
        
        if "Peak" in line and "Partial" not in line and "Off" not in line:
            result["C"] = float(nums[0].replace(",", ""))
            if len(nums) > 1: result["F"] = float(nums[-1].replace(",", ""))
        elif "Partial Peak" in line:
            result["D"] = float(nums[0].replace(",", ""))
            if len(nums) > 1: result["G"] = float(nums[-1].replace(",", ""))
        elif "Off Peak" in line:
            result["E"] = float(nums[0].replace(",", ""))

    # 2. กลุ่ม Energy (I, J, K) - ดึงค่าสุดท้ายของบรรทัดที่มี P, PP, OP
    for line in text.split('\n'):
        nums = re.findall(r"[\d,]+\.\d+", line)
        if not nums: continue
        # บรรทัดหน่วยมักมีคำว่า 'พลังงานไฟฟ้า' หรือตัวย่อ
        if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line:
            result["I"] = float(nums[-1].replace(",", ""))
        elif "PP" in line:
            result["J"] = float(nums[-1].replace(",", ""))
        elif "OP" in line:
            result["K"] = float(nums[-1].replace(",", ""))

    # 3. ข้อมูลอื่นๆ (ดึงด้วย Keyword)
    def get_val(pattern):
        m = re.search(pattern, text, re.I)
        return float(m.group(1).replace(",", "")) if m else ""

    result["L"] = get_val(r'เงินค่าไฟฟ้าฐาน.*?([\d,]+\.\d+)')
    result["M"] = get_val(r'ค่า Ft.*?([\d,]+\.\d+)')
    result["N"] = get_val(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)')
    result["P"] = get_val(r'คาเพาเวอร์แฟคเตอร.*?([\d,]+\.\d+)')
    result["Q"] = get_val(r'รวมเงินค่าไฟฟ้า \(Sub Total\).*?([\d,]+\.\d+)')

    return result

# ส่วน UI
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", accept_multiple_files=True)
if uploaded_files:
    results = [extract_exact_pea_bill(f) for f in uploaded_files]
    df = pd.DataFrame(results)
    edited_df = st.data_editor(df, use_container_width=True)
    
    # ดาวน์โหลด
    output = BytesIO()
    edited_df.to_excel(output, index=False)
    st.download_button("🟢 ดาวน์โหลด Excel", data=output.getvalue(), file_name="PEA_Report.xlsx")
