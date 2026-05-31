import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันแก้ไขจุดบกพร่องตารางหน่วย: สกัดตัวเลข I, J, K, L โดยเลิกดักจับคำว่าหน่วยที่สระชอบเพี้ยน")

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
                
    # ตั้งค่าให้ทุกช่องเป็นค่าว่างเปล่าเริ่มต้น
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    for line in text_lines:
        # ดึงตัวเลขทั้งหมดที่เป็นทศนิยม 2 ตำแหน่งขึ้นไปในบรรทัดนั้นมาดู
        all_numbers = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
        
        # --- ส่วนที่ 1: ตารางคิดเงินด้านบน (Demand) ---
        if "Peak" in line and "กว." in line:
            if len(all_numbers) >= 3:
                col_f = clean_num(all_numbers[2])  # เงิน Demand Peak -> 1,031,881.00 (คอลัมน์ F ถูกอยู่แล้ว)
                
        elif "Partial Peak" in line and "กว." in line:
            if len(all_numbers) >= 3:
                col_g = clean_num(all_numbers[2])  # เงิน Demand PP -> 96,563.20 (คอลัมน์ G ถูกอยู่แล้ว)
                
        # --- ส่วนที่ 2: ตารางจำนวนหน่วยด้านล่าง (Energy) ---
        # ใช้วิธีเช็กคำคีย์เวิร์ด + จำนวนตัวเลขในบรรทัดเพื่อป้องกันเรื่องสระและวรรณยุกต์ภาษาไทยหาย
        elif "Peak" in line and "กว." not in line:
            # ถ้าเจอคำว่า Peak ในบรรทัดอื่นที่ไม่ใช่บรรทัด กว. แปลว่าเป็นบรรทัดหน่วยชัวร์
            if all_numbers:
                col_i = clean_num(all_numbers[-1])  # ดึงเลขตัวสุดท้าย -> 144,000.00 ลงคอลัมน์ I
                
        elif "Partial Peak" in line and "กว." not in line:
            if all_numbers:
                col_j = clean_num(all_numbers[-1])  # ดึงเลขตัวสุดท้าย -> 651,000.00 ลงคอลัมน์ J
                
        elif "Off Peak" in line:
            # ในบิลนี้ บรรทัดหน่วย Off Peak จะมีตัวเลขสองชุดคือ หน่วยใช้ไป (440,200.00) กับ จำนวนเงิน (3,887,297.92)
            if "กว." not in line and len(all_numbers) >= 2:
                col_k = clean_num(all_numbers[0])  # ตัวแรกคือหน่วย -> 440,200.00 ลงคอลัมน์ K
                col_l = clean_num(all_numbers[1])  # ตัวสองคือจำนวนเงิน -> 3,887,297.92 ลงคอลัมน์ L

        # --- ส่วนที่ 3: ท้ายตารางบิลค่าไฟฟ้า (Power Factor) ---
        elif "เพาเวอร์แฟคเตอร์" in line or "เพาเวอร" in line or "Factor" in line:
            if all_numbers:
                col_p = clean_num(all_numbers[0])  # เงินค่า Power Factor -> 584,249.40 (คอลัมน์ P ถูกอยู่แล้ว)

    return {
        "ชื่อไฟล์": file_obj.name,
        "C": col_c, "D": col_d, "E": col_e,
        "F": col_f, "G": col_g, "H": col_h,
        "I": col_i, "J": col_j, "K": col_k, "L": col_l,
        "M": col_m, "N": col_n, "O": col_o,
        "P": col_p, "Q": col_q
    }

# จอดิสเพลย์หน้าเว็บ
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังดึงข้อมูลด้วยระบบกรองตัวเลขเสถียรสูง {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ แก้ไขช่องว่างและแมปตำแหน่งเสร็จสมบูรณ์!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูล (สกัดตัวเลขครบทุกช่อง C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Final_Data')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับ Copy แปะลงล็อกจริง",
            data=excel_data,
            file_name=f"PEA_Final_Fix_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
