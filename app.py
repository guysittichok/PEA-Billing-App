import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบสกัดข้อมูลบิล PEA", layout="wide")
st.title("⚡ ระบบสกัดข้อมูลบิลค่าไฟฟ้า PEA (คัดเฉพาะข้อมูลที่แม่นยำ)")

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    # เหลือเฉพาะช่องที่ดึงแล้วชัวร์และตรงล็อก (ตัด M, N, O, Q ออก)
    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", 
        "I": "", "J": "", "K": "", "L": "", "P": ""
    }

    # 1. Demand Charge (C, D, E, F, G)
    for line in text.split('\n'):
        nums = re.findall(r"[\d,]+\.\d+", line)
        if not nums: continue
        if "Peak" in line and "Partial" not in line and "Off" not in line:
            result["C"] = float(nums[0].replace(",", "")); result["F"] = float(nums[-1].replace(",", ""))
        elif "Partial Peak" in line:
            result["D"] = float(nums[0].replace(",", "")); result["G"] = float(nums[-1].replace(",", ""))
        elif "Off Peak" in line:
            result["E"] = float(nums[0].replace(",", ""))

    # 2. Energy Usage (I, J, K) - ดึงค่าสุดท้ายของบรรทัด
    for line in text.split('\n'):
        nums = re.findall(r"[\d,]+\.\d+", line)
        if not nums: continue
        if ("Peak" in line and "P" in line and "พลังงาน" in line) or ("Peak" in line and "P" in line and "Unit" in line):
            result["I"] = float(nums[-1].replace(",", ""))
        elif "Partial" in line and "PP" in line:
            result["J"] = float(nums[-1].replace(",", ""))
        elif "Off" in line and "OP" in line:
            result["K"] = float(nums[-1].replace(",", ""))

    # 3. เงินค่าไฟฟ้าฐาน (L) และ Power Factor (P)
    # ใช้ Regex ที่ดึงเลขบรรทัดที่มีคีย์เวิร์ดชัดเจนเท่านั้น
    l_match = re.search(r'เงินค่าไฟฟ้าฐาน.*?([\d,]+\.\d+)', text)
    if l_match: result["L"] = float(l_match.group(1).replace(",", ""))
    
    p_match = re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I)
    if p_match: result["P"] = float(p_match.group(1).replace(",", ""))

    return result

# Streamlit UI
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    data = [extract_exact_pea_bill(f) for f in uploaded_files]
    df = pd.DataFrame(data)
    
    # แสดงผลเฉพาะช่องที่เลือกไว้ (C ถึง L และ P)
    edited_df = st.data_editor(df, use_container_width=True)
    
    # ดาวน์โหลด Excel
    output = BytesIO()
    edited_df.to_excel(output, index=False)
    st.download_button("🟢 ดาวน์โหลด Excel เฉพาะข้อมูลที่แม่นยำ", data=output.getvalue(), file_name="PEA_Clean_Report.xlsx")
