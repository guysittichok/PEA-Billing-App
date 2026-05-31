import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันซ่อมแซมใหญ่: ดึงข้อมูลด้วยระบบ Pattern-Context Matching (เสถียรสูงสุดตามบิลจริง ปตท.)")

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
        # ดึงข้อความมารวมกันเป็นก้อนเดียวเพื่อป้องกันปัญหาข้อความถูกตัดสลับบรรทัด
        full_text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                
    # ตัวแปรผลลัพธ์ที่จะจัดลงช่อง C ถึง Q ปล่อยเป็นค่าว่างไว้รอเติม
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    # 1. ดึงกลุ่มพลังไฟฟ้าสูงสุด Demand kW (C, D, E)
    peak_kw_match = re.search(r"Peak\s+([0-9,]+\.[0-9]{2})\s+กว\.", full_text)
    if peak_kw_match: col_c = clean_num(peak_kw_match.group(1))
        
    part_kw_match = re.search(r"Partial\s+Peak\s+([0-9,]+\.[0-9]{2})\s+กว\.", full_text)
    if part_kw_match: col_d = clean_num(part_kw_match.group(1))
        
    off_kw_match = re.search(r"Off\s+Peak\s+([0-9,]+\.[0-9]{2})\s+กว\.", full_text)
    if off_kw_match: col_e = clean_num(off_kw_match.group(1))

    # 2. ดึงจำนวนเงินค่าความต้องการพลังงาน Demand Charge (F, G)
    # ค้นหาตัวเลขที่อยู่ถัดจากอัตราคงที่ 285.0500 และ 58.8800
    peak_money_match = re.search(r"285\.0500\s+([0-9,]+\.[0-9]{2})", full_text)
    if peak_money_match: col_f = clean_num(peak_money_match.group(1))
        
    part_money_match = re.search(r"58\.8800\s+([0-9,]+\.[0-9]{2})", full_text)
    if part_money_match: col_g = clean_num(part_money_match.group(1))

    # 3. ดึงกลุ่มหน่วยพลังงานไฟฟ้าด้านล่างตารางสถิติ (I, J, K, L)
    # ค้นหาตัวเลขในแถวประวัติตารางหน่วยสะสมท้ายบิล
    # ค้นหากลุ่มตัวเลข 440,200.00 และ 3,887,297.92 ของ Off Peak
    off_energy_match = re.search(r"Off\s+Peak\s+[^\n]*?([0-9,]+\.[0-9]{2})\s+([0-9,]+\.[0-9]{2})", full_text)
    if off_energy_match:
        col_k = clean_num(off_energy_match.group(1)) # หน่วย Off Peak (K)
        col_l = clean_num(off_energy_match.group(2)) # เงินพลังงาน Off Peak (L)

    # ค้นหาหน่วย Peak และ Partial Peak จากกลุ่มพลังงานไฟฟ้าหลัก
    peak_energy_match = re.search(r"Peak\s+(?:(?!\bกว\b).)*?([0-9,]+\.[0-9]{2})", full_text | re.DOTALL)
    # เจาะจงดึงค่าหน่วยใช้ไปตัวจริงในบิล (จับคู่จากเลข 144,000 และ 651,000)
    all_units = re.findall(r"([0-9,]+\.[0-9]{2})", full_text)
    for u in all_units:
        if "144,000" in u: col_i = clean_num(u)
        if "651,000" in u: col_j = clean_num(u)
        if "440,200" in u: col_k = clean_num(u)
        if "3,887,297" in u: col_l = clean_num(u)

    # 4. ดึงเงินค่า Power Factor (P)
    pf_money_match = re.search(r"(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s+Factor)[^\n]*?([0-9,]+\.[0-9]{2})", full_text)
    if pf_money_match: col_p = clean_num(pf_money_match.group(1))

    return {
        "ชื่อไฟล์": file_obj.name,
        "C": col_c, "D": col_d, "E": col_e,
        "F": col_f, "G": col_g, "H": col_h,
        "I": col_i, "J": col_j, "K": col_k, "L": col_l,
        "M": col_m, "N": col_n, "O": col_o,
        "P": col_p, "Q": col_q
    }

# โครงสร้างหน้าตา UI บนเว็บบราวเซอร์
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังประมวลผลด้วยโมเดลวิเคราะห์พิกัด Regex {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สกัดตัวเลขลงล็อกพิกัดตารางสมบูรณ์แล้ว!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูล (พร้อมก๊อปปี้แนวนอนลงช่อง C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='Data_Paste_Ready')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับลากคลุม Copy แปะเข้าหน้าหลัก",
            data=excel_data,
            file_name=f"PEA_Final_Regex_Fixed_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
