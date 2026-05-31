import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันแก้ไขสมบูรณ์: ใช้ระบบค้นหาเชิงโครงสร้างตัวเลข ป้องกันปัญหาภาษาไทยสะกดเพี้ยนและข้ามบรรทัด")

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
                
    # ตัวแปรผลลัพธ์ C ถึง Q สำหรับนำไปใช้แปะตารางหลัก
    col_f = col_g = col_i = col_j = col_k = col_l = col_p = ""
    
    # ดึงกลุ่มตัวเลขทั้งหมดที่มีในแต่ละบรรทัดเก็บไว้ล่วงหน้าเพื่อทำดัชนีข้อมูล
    for line in text_lines:
        # ล้างช่องว่างส่วนเกินและค้นหากลุ่มตัวเลขทศนิยม
        line_clean = line.strip()
        all_numbers = re.findall(r"([0-9,]+\.[0-9]{2,4})", line_clean)
        
        # 1. สกัดกลุ่มเงิน Demand Charge (F และ G)
        # มองหาอัตราคงที่ 285.05 (Peak Demand Rate) เพื่อดึงเงินช่อง F
        if "285.05" in line_clean and len(all_numbers) >= 2:
            col_f = clean_num(all_numbers[-1])
        # มองหาอัตราคงที่ 58.88 (Partial Peak Demand Rate) เพื่อดึงเงินช่อง G
        elif "58.88" in line_clean and len(all_numbers) >= 2:
            col_g = clean_num(all_numbers[-1])
            
        # 2. สกัดกลุ่มหน่วยพลังงานไฟฟ้าและค่าไฟฟ้า (I, J, K, L)
        # ตรวจสอบจากส่วนสรุปจำนวนหน่วยสะสมที่ใช้จริงในรอบเดือน
        if "Peak" in line_clean and "กว." not in line_clean and "285.05" not in line_clean:
            if all_numbers:
                col_i = clean_num(all_numbers[-1]) # หน่วย Peak (I)
        elif "Partial Peak" in line_clean and "กว." not in line_clean and "58.88" not in line_clean:
            if all_numbers:
                col_j = clean_num(all_numbers[-1]) # หน่วย Partial Peak (J)
        elif "Off Peak" in line_clean and "กว." not in line_clean:
            # บรรทัด Off Peak พลังงานไฟฟ้าจะมีตัวเลข 2 ชุดเสมอ คือ จำนวนหน่วย และ จำนวนเงิน
            if len(all_numbers) >= 2:
                col_k = clean_num(all_numbers[0])  # หน่วย Off Peak (K)
                col_l = clean_num(all_numbers[1])  # เงินพลังงาน Off Peak (L)

        # 3. สกัดกลุ่มเงินค่า Power Factor (P)
        if "เพาเวอร์" in line_clean or "แฟคเตอร์" in line_clean or "Factor" in line_clean:
            if all_numbers:
                col_p = clean_num(all_numbers[0])   # เงินค่า Power Factor (P)

    # ประกอบร่างคืนค่ากลับไปเป็นหน้ากระดานกว้างล็อกพิกัด C ถึง Q
    return {
        "ชื่อไฟล์": file_obj.name,
        "C": "", "D": "", "E": "",
        "F": col_f,
        "G": col_g,
        "H": "",
        "I": col_i,
        "J": col_j,
        "K": col_k,
        "L": col_l,
        "M": "", "N": "", "O": "",
        "P": col_p,
        "Q": ""
    }

# ฟังก์ชันหน้าตา UI บนเว็บบราวเซอร์
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสกัดข้อมูลระบบ Pattern-Matching {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ ประมวลผลและสกัดข้อมูลลงพิกัดสำเร็จ!")
        st.subheader("📊 2. ตารางพรีวิวก่อน Copy แปะลงช่อง C ถึง Q")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Data_Extract')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์รายงานสำเร็จรูป",
            data=excel_data,
            file_name=f"PEA_Final_Structure_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
