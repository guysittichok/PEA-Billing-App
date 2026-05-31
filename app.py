import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันทำงานแบบเทกระจาด (Raw Data Extraction): ดึงตัวเลขทั้งหมดแล้วจัดเรียงตามดัชนีโครงสร้างบิล")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def clean_num(val_str):
    if not val_str: return ""
    return float(val_str.replace(",", "").strip())

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        # ดึงข้อความมารวมกันเป็นก้อนเดียว
        full_text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                
    # ตัวแปรผลลัพธ์เริ่มต้นสำหรับ Column C ถึง Q
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    # ดึงตัวเลขทั้งหมดในบิลที่มีทศนิยม 2-4 ตำแหน่ง ออกมารวมกันเป็นลิสต์ยาว (เทกระจาดข้อมูล)
    all_numbers = re.findall(r"([0-9,]+\.[0-9]{2,4})", full_text)
    
    # ตรวจสอบและแมปข้อมูลตามลำดับที่ปรากฏในโครงสร้างบิลมาตรฐานรายใหญ่ (4.1.2.4)
    if len(all_numbers) >= 17:
        col_c = clean_num(all_numbers[0])   # 3,620.00  -> ช่อง C (Peak kW)
        col_f = clean_num(all_numbers[2])   # 1,031,881.00 -> ช่อง F (เงิน Demand Peak)
        
        col_d = clean_num(all_numbers[3])   # 5,260.00  -> ช่อง D (Partial Peak kW)
        col_g = clean_num(all_numbers[5])   # 96,563.20 -> ช่อง G (เงิน Demand PP)
        
        col_e = clean_num(all_numbers[6])   # 5,240.00  -> ช่อง E (Off Peak kW)
        
        col_i = clean_num(all_numbers[9])   # 144,000.00 -> ช่อง I (หน่วย Peak)
        col_j = clean_num(all_numbers[10])  # 651,000.00 -> ช่อง J (หน่วย Partial Peak)
        col_k = clean_num(all_numbers[11])  # 440,200.00 -> ช่อง K (หน่วย Off Peak)
        col_l = clean_num(all_numbers[12])  # 3,887,297.92 -> ช่อง L (เงินพลังงาน Off Peak)
        
        col_p = clean_num(all_numbers[16])  # 584,249.40 -> ช่อง P (เงินค่า Power Factor)

    return {
        "ชื่อไฟล์": file_obj.name,
        "C": col_c, "D": col_d, "E": col_e,
        "F": col_f, "G": col_g, "H": col_h,
        "I": col_i, "J": col_j, "K": col_k, "L": col_l,
        "M": col_m, "N": col_n, "O": col_o,
        "P": col_p, "Q": col_q
    }

# หน้าจอการทำงานหลักของเว็บแอป Streamlit
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังดูดข้อมูลแบบเทกระจาดตามดัชนีโครงสร้าง... {f.name}"):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ เรียงช่องข้อมูลด้วยระบบ Indexing เรียบร้อย!")
        st.subheader("📊 2. ตารางพรีวิวก่อนคัดลอก (เรียงช่อง C ถึง Q ตามผัง Excel หลัก)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Raw_Index_Map')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับก๊อปปี้ไปแปะในหน้าหลัก",
            data=excel_data,
            file_name=f"PEA_Raw_Extracted_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
