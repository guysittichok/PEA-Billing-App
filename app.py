import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันแก้ไขสมบูรณ์: สกัดตามผังตำแหน่ง C-Q ของจริง (คัดแยกกลุ่ม Demand, Energy และค่าบริการอย่างแม่นยำ)")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # ดีฟอลต์เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def clean_num(val_str):
    if not val_str: return 0.0
    return float(val_str.replace(",", "").strip())

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        text_lines = text.split('\n')
                
    # ตัวแปรตั้งต้นสำหรับคอลลัมน์ C - Q
    col_c = col_d = col_e = col_f = col_g = col_i = col_j = col_k = col_l = col_n = col_p = 0.0
    
    for line in text_lines:
        # 1. จัดการกลุ่ม Peak (ดึง kW ลงช่อง C, ดึงจำนวนเงินลงช่อง F)
        if "Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_c = clean_num(nums[0])  # Peak kW -> 3,620.00
                col_f = clean_num(nums[2])  # จำนวนเงิน Peak -> 1,031,881.00
                
        # 2. จัดการกลุ่ม Partial Peak (ดึง kW ลงช่อง D, ดึงจำนวนเงินลงช่อง G)
        elif "Partial Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_d = clean_num(nums[0])  # Partial Peak kW -> 5,260.00
                col_g = clean_num(nums[2])  # จำนวนเงิน Partial Peak -> 96,563.20
                
        # 3. จัดการกลุ่ม Off Peak (ดึง kW ลงช่อง E)
        elif "Off Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums:
                col_e = clean_num(nums[0])  # Off Peak kW -> 5,240.00
                
        # 4. จัดการกลุ่มหน่วยพลังงานไฟฟ้าด้านล่าง (ดึงหน่วยลงช่อง I, J, K และดึงจำนวนเงินลงช่อง L)
        elif "Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums:
                col_i = clean_num(nums[-1])  # หน่วย Peak -> 144,000.00
        elif "Partial Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums:
                col_j = clean_num(nums[-1])  # หน่วย PP -> 651,000.00
        elif "Off Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 2:
                col_k = clean_num(nums[0])  # หน่วย Off Peak -> 440,200.00
                col_l = clean_num(nums[1])  # จำนวนเงินของ Off Peak -> 3,887,297.92

        # 5. จัดการกลุ่มค่าบริการ และค่า Power Factor ท้ายตารางเงิน (ลงช่อง N และ P)
        elif "ค่าบริการ" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: 
                col_n = clean_num(num.group(1))  # ค่าบริการ -> 312.24
        elif "คาเพาเวอร์แฟคเตอร" in line or "เพาเวอร์แฟคเตอร์" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: 
                col_p = clean_num(num.group(1))  # ค่า Power Factor -> 584,249.40

    # ประกอบโครงสร้างข้อมูลส่งออกแบบเว้นช่องว่างคั่นตามล็อกช่องจริงในฟอร์แมต Excel (C ถึง Q)
    return {
        "ชื่อไฟล์": file_obj.name,
        "C: Peak (kW)": col_c,
        "D: Partial Peak (kW)": col_d,
        "E: Off Peak (kW)": col_e,
        "F: [เงินบรรทัด Peak]": col_f,
        "G: [เงินบรรทัด PP]": col_g,
        "H: (ว่าง)": "",
        "I: [หน่วย Peak]": col_i,
        "J: [หน่วย PP]": col_j,
        "K: [หน่วย Off Peak]": col_k,
        "L: [เงินบรรทัด Off Peak]": col_l,
        "M: (ว่าง)": "",
        "N: ค่าบริการ (Baht)": col_n,
        "O: (ว่าง)": "",
        "P: ค่า Power Factor (Baht)": col_p,
        "Q: (ว่าง)": ""
    }

st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังวิเคราะห์แมปตัวเลขไฟล์ {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สกัดข้อมูลลงตำแหน่งช่องตารางสำเร็จแล้ว!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูลเพื่อคัดลอกแนวนอน (C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Final_Report')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับ Copy แปะรวดเดียวพอดีช่อง",
            data=excel_data,
            file_name=f"PEA_Ready_To_Paste_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
