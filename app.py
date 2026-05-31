import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันสมบูรณ์ที่สุด: จัดล็อกตำแหน่งและเว้นช่องว่างตามผังบัญชีบิลรายใหญ่ (คอลัมน์ C ถึง Q)")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def clean_num(val_str):
    if not val_str: return 0.0
    return float(val_str.replace(",", "").strip())

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        text_lines = text.split('\n')
                
    # ตัวแปรตั้งต้นตามคอลัมน์ที่คุณมาร์คไว้
    col_c = col_d = col_e = col_f = col_g = col_i = col_j = col_k = col_l = col_n = col_p = 0.0
    
    for line in text_lines:
        # กลุ่ม Peak
        if "Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_c = clean_num(nums[0]) # Peak (kW)
                col_f = clean_num(nums[2]) # เงิน Demand Peak
            elif "Peak" in line and "หนวย" in line:
                nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
                if nums_kwh: col_i = clean_num(nums_kwh[-1]) # หน่วย Peak

        # กลุ่ม Partial Peak
        elif "Partial Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_d = clean_num(nums[0]) # PP (kW)
                col_g = clean_num(nums[2]) # เงิน Demand PP
            elif "Partial Peak" in line and "หนวย" in line:
                nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
                if nums_kwh: col_j = clean_num(nums_kwh[-1]) # หน่วย PP

        # กลุ่ม Off Peak
        elif "Off Peak" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if "กว." in line and len(nums) >= 2:
                col_e = clean_num(nums[0]) # Off Peak (kW)
            elif "หนวย" in line and len(nums) >= 2:
                col_k = clean_num(nums[0]) # หน่วย Off Peak
                col_l = clean_num(nums[1]) # เงินพลังงาน Off Peak

        # กลุ่มค่าใช้จ่ายอื่นๆ
        elif "ค่า Ft" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: col_n = clean_num(num.group(1))
        elif "คาเพาเวอร์แฟคเตอร" in line or "เพาเวอร์แฟคเตอร์" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: col_p = clean_num(num.group(1))

    # ล็อกและเรียงคอลัมน์แนวนอนให้ตรงกับไฟล์หลักของคุณเป๊ะๆ (ช่องที่ไม่ใช้กรอก ปล่อยว่างไว้)
    return {
        "ชื่อไฟล์": file_obj.name,
        "C: Peak (kW)": col_c,
        "D: Partial Peak (kW)": col_d,
        "E: Off Peak (kW)": col_e,
        "F: [Demand เงิน Peak]": col_f,
        "G: [Demand เงิน PP]": col_g,
        "H: (ปล่อยว่าง)": "",
        "I: [หน่วย Peak]": col_i,
        "J: [หน่วย PP]": col_j,
        "K: [หน่วย Off Peak]": col_k,
        "L: [พลังงาน เงิน Off Peak]": col_l,
        "M: (ปล่อยว่าง)": "",
        "N: ค่า Ft (Baht)": col_n,
        "O: (ปล่อยว่าง)": "",
        "P: ค่า Power Factor (Baht)": col_p,
        "Q: (ปล่อยว่าง)": ""
    }

st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังวิเคราะห์ไฟล์ {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สกัดข้อมูลจัดตำแหน่งเรียบร้อย!")
        st.subheader("📊 2. ตารางรวมข้อมูลเพื่อใช้ Copy (C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Matrix_Report')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับก๊อปปี้ข้อมูล",
            data=excel_data,
            file_name=f"Billing_PEA_Matrix_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
