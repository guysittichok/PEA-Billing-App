import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("⚡ ระบบสกัดข้อมูลบิล PEA (รองรับทุกประเภทบิล)")

def extract_pea_data(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        # ดึงข้อความออกมาเป็นก้อนเดียว
        full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])

    # สร้าง Dictionary เปล่า
    data = {k: "" for k in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]}
    data["ชื่อไฟล์"] = file_obj.name

    # Helper function: ค้นหาตัวเลขที่อยู่ใกล้คำหลักที่สุด
    def find_val(anchor, line_text):
        # ค้นหาบรรทัดที่มี Anchor อยู่
        for line in line_text.split('\n'):
            if anchor in line:
                nums = re.findall(r"[\d,]+\.\d+", line)
                return float(nums[-1].replace(",", "")) if nums else ""
        return ""

    # ดึงข้อมูล Demand (C, D, E) และ เงิน (F, G)
    # ใช้การค้นหาแบบอิสระ ไม่ยึดติดรูปแบบบรรทัด
    data["C"] = find_val("Peak", full_text) if "Peak" in full_text else ""
    data["D"] = find_val("Partial Peak", full_text) if "Partial Peak" in full_text else ""
    data["E"] = find_val("Off Peak", full_text) if "Off Peak" in full_text else ""
    
    # ดึงเงินค่าไฟฟ้าฐาน (L) - ใช้คำค้นหาที่แน่นอน
    data["L"] = find_val("เงินค่าไฟฟ้าฐาน", full_text)
    
    # ดึง Ft และ Power Factor
    data["M"] = find_val("ค่า Ft", full_text)
    data["P"] = find_val("คาเพาเวอร์แฟคเตอร", full_text)
    data["Q"] = find_val("รวมเงินค่าไฟฟ้า (Sub Total)", full_text)

    return data

# หน้าจอหลัก
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", accept_multiple_files=True)
if uploaded_files:
    results = [extract_pea_data(f) for f in uploaded_files]
    df = pd.DataFrame(results)
    
    st.write("### ตารางข้อมูลที่สกัดได้")
    edited_df = st.data_editor(df, use_container_width=True)
    
    # ปุ่มดาวน์โหลด
    output = BytesIO()
    edited_df.to_excel(output, index=False)
    st.download_button("🟢 ดาวน์โหลด Excel", data=output.getvalue(), file_name="PEA_Data.xlsx")
