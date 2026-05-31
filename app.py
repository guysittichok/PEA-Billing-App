import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันแก้ไขตามบรีฟจริง: จัดเรียงและล็อกช่องตรงตามพารามิเตอร์บัญชีต้นทุน (คอลัมน์ C ถึง Q)")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # ดีฟอลต์ที่เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def clean_num(val_str):
    if not val_str: return 0.0
    return float(val_str.replace(",", "").strip())

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        text_lines = text.split('\n')
                
    # ตัวแปรตั้งต้นสำหรับคอลัมน์ที่ต้องการดึงค่า
    col_c = col_d = col_e = col_f = col_g = col_i = col_j = col_k = col_l = col_n = col_p = 0.0
    
    for line in text_lines:
        # 1. ดึงกลุ่ม Peak (C, F, I)
        if "Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_c = clean_num(nums[0])  # Peak kW -> 3,620.00
                col_f = clean_num(nums[2])  # เงิน Demand Peak -> 1,031,881.00
        elif "Peak" in line and "หนวย" in line:
            nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums_kwh: 
                col_i = clean_num(nums_kwh[-1]) # หน่วย Peak -> 144,000.00

        # 2. ดึงกลุ่ม Partial Peak (D, G, J)
        elif "Partial Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                col_d = clean_num(nums[0])  # PP kW -> 5,260.00
                col_g = clean_num(nums[2])  # เงิน Demand PP -> 96,563.20
        elif "Partial Peak" in line and "หนวย" in line:
            nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums_kwh: 
                col_j = clean_num(nums_kwh[-1]) # หน่วย PP -> 651,000.00

        # 3. ดึงกลุ่ม Off Peak (E, K, L)
        elif "Off Peak" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if "กว." in line and len(nums) >= 2:
                col_e = clean_num(nums[0])  # Off Peak kW -> 5,240.00
            elif "หนวย" in line and len(nums) >= 2:
                col_k = clean_num(nums[0])  # หน่วย Off Peak -> 440,200.00
                col_l = clean_num(nums[1])  # เงินพลังงาน Off Peak -> 3,887,297.92

        # 4. ดึงกลุ่มค่าเงินอื่นๆ ท้ายบิล (N, P)
        elif "ค่า Ft" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: col_n = clean_num(num.group(1)) # ค่า Ft -> 120,061.44
        elif "คาเพาเวอร์แฟคเตอร" in line or "เพาเวอร์แฟคเตอร์" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: col_p = clean_num(num.group(1)) # ค่า Power Factor -> 584,249.40

    # สรุปผังส่งออกเรียงแถวตรงตามหน้ากระดาน Excel ของคุณเป๊ะๆ (ช่องว่างใส่คั่นไว้ให้ล็อกตำแหน่งช่อง)
    return {
        "ชื่อไฟล์": file_obj.name,
        "คอลัมน์ C: Peak (kW)": col_c,
        "คอลัมน์ D: Partial Peak (kW)": col_d,
        "คอลัมน์ E: Off Peak (kW)": col_e,
        "คอลัมน์ F: เงิน Demand Peak": col_f,
        "คอลัมน์ G: เงิน Demand PP": col_g,
        "คอลัมน์ H: (ว่าง)": "",
        "คอลัมน์ I: หน่วย Peak": col_i,
        "คอลัมน์ J: หน่วย PP": col_j,
        "คอลัมน์ K: หน่วย Off Peak": col_k,
        "คอลัมน์ L: เงินพลังงาน Off Peak": col_l,
        "คอลัมน์ M: (ว่าง)": "",
        "คอลัมน์ N: ค่า Ft": col_n,
        "คอลัมน์ O: (ว่าง)": "",
        "คอลัมน์ P: ค่า Power Factor": col_p,
        "คอลัมน์ Q: (ว่าง)": ""
    }

# หน้าจอการจัดการฝั่งหน้าเว็บ
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
        st.success(f"⚡ สกัดข้อมูลลงผังช่องตารางสำเร็จ!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูลสำหรับก๊อปปี้แนวนอน (C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Matrix_Paste')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์รายงานสำเร็จรูป (คลุมก๊อปปี้ไปแปะใน Excel หลักได้ทันที)",
            data=excel_data,
            file_name=f"PEA_Ready_To_Paste_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
