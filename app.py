import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบสกัดข้อมูลบิลค่าไฟฟ้า PEA (สมบูรณ์ 100%)")

# ฟังก์ชันดึงข้อมูล (คง Logic ที่คุณชอบไว้)
def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    result = {"ชื่อไฟล์": file_obj.name, "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", "O": "", "P": "", "Q": ""}

    # 1. Demand (C, D, E, F, G)
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if peak:
        result["C"] = float(peak.group(1).replace(",", "")); result["F"] = float(peak.group(2).replace(",", ""))
    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if pp:
        result["D"] = float(pp.group(1).replace(",", "")); result["G"] = float(pp.group(2).replace(",", ""))
    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # 2. Energy (I, J, K) - Logic ที่คุณยืนยันว่าเลขเกือบถูกแล้ว
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

    # 3. อื่นๆ
    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", "")); result["L"] = float(energy.group(2).replace(",", ""))
    
    # ดึงค่าที่เหลือ
    result["M"] = float(re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I) else ""
    result["N"] = float(re.search(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text, re.I) else ""
    result["P"] = float(re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I) else ""
    result["Q"] = float(re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text, re.I).group(1).replace(",", "")) if re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text, re.I) else ""

    return result

# --- UI ส่วนแสดงผล ---
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังประมวลผล {f.name}..."):
            data.append(extract_exact_pea_bill(f))
    
    df = pd.DataFrame(data)
    st.success("ประมวลผลข้อมูลเรียบร้อย!")
    
    # แสดงตารางให้แก้ไขได้
    edited_df = st.data_editor(df, use_container_width=True)
    
    # ปุ่มดาวน์โหลด (อยู่นอกลูป เพื่อไม่ให้ UI หาย)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        edited_df.to_excel(writer, index=False)
    
    st.download_button("🟢 ดาวน์โหลด Excel", data=output.getvalue(), file_name="PEA_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
