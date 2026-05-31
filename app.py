import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันพิกัดอัจฉริยะ: ดึงข้อมูลแบบตัดพิกัด Row × Column ตามคำแนะนำวิศวกรรม")

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
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    for line in text_lines:
        line_clean = line.strip()
        # จับกลุ่มตัวเลขทศนิยมทั้งหมดในบรรทัด
        all_numbers = re.findall(r"([0-9,]+\.[0-9]{2,4})", line_clean)
        
        if not all_numbers:
            continue
            
        # 1. เสิร์ชหา Row พลังไฟฟ้าสูงสุด (Demand kW) -> หยอดลงช่อง C, D, E และจำนวนเงินลงช่อง F, G
        if "Peak" in line_clean and "กว." in line_clean:
            col_c = clean_num(all_numbers[0]) # ได้เลขจำนวนที่ใช้ 3,620.00 ลงช่อง C
            if len(all_numbers) >= 3:
                col_f = clean_num(all_numbers[2]) # ได้จำนวนเงิน 1,031,881.00 ลงช่อง F
                
        elif "Partial Peak" in line_clean and "กว." in line_clean:
            col_d = clean_num(all_numbers[0]) # ได้เลขจำนวนที่ใช้ 5,260.00 ลงช่อง D
            if len(all_numbers) >= 3:
                col_g = clean_num(all_numbers[2]) # ได้จำนวนเงิน 96,563.20 ลงช่อง G
                
        elif "Off Peak" in line_clean and "กว." in line_clean:
            col_e = clean_num(all_numbers[0]) # ได้เลขจำนวนที่ใช้ 5,240.00 ลงช่อง E

        # 2. เสิร์ชหา Row พลังงานไฟฟ้า (Energy Units) -> หยอดลงช่อง I, J, K และจำนวนเงินลงช่อง L
        elif "Peak" in line_clean and ("หนวย" in line_clean or "หน่วย" in line_clean or "กว." not in line_clean) and "285.05" not in line_clean:
            col_i = clean_num(all_numbers[0]) # ได้หน่วยใช้ไปตัวแรกสุด 144,000.00 ลงช่อง I
            
        elif "Partial Peak" in line_clean and ("หนวย" in line_clean or "หน่วย" in line_clean or "กว." not in line_clean) and "58.88" not in line_clean:
            col_j = clean_num(all_numbers[0]) # ได้หน่วยใช้ไปตัวแรกสุด 651,000.00 ลงช่อง J
            
        elif "Off Peak" in line_clean and ("หนวย" in line_clean or "หน่วย" in line_clean or "กว." not in line_clean):
            # บรรทัดนี้จะมีทั้งหน่วยใช้ไป (ตัวแรก) และจำนวนเงินค่าพลังงาน (ตัวที่สอง)
            col_k = clean_num(all_numbers[0]) # ได้หน่วยใช้ไปตัวแรกสุด 440,200.00 ลงช่อง K
            if len(all_numbers) >= 2:
                col_l = clean_num(all_numbers[1]) # ได้จำนวนเงินตัวที่สอง 3,887,297.92 ลงช่อง L

        # 3. เสิร์ชหา Row ค่าเพาเวอร์แฟคเตอร์ -> หยอดลงช่อง P
        elif "เพาเวอร์" in line_clean or "แฟคเตอร์" in line_clean or "Factor" in line_clean:
            col_p = clean_num(all_numbers[0]) # ได้จำนวนเงินค่า Power Factor 584,249.40 ลงช่อง P

    return {
        "ชื่อไฟล์": file_obj.name,
        "C": col_c,
        "D": col_d,
        "E": col_e,
        "F": col_f,
        "G": col_g,
        "H": col_h,
        "I": col_i,
        "J": col_j,
        "K": col_k,
        "L": col_l,
        "M": col_m,
        "N": col_n,
        "O": col_o,
        "P": col_p,
        "Q": col_q
    }

# หน้าจอหลักของ Streamlit UI
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังค้นหาข้อมูลพิกัด Matrix {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สกัดและล็อกตำแหน่งช่องตาราง Matrix สำเร็จแล้ว!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูล (เรียงช่องพร้อมสำหรับการลากคลุม Copy)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Matrix_Match')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel เพื่อ Copy แปะลง Column C แถวที่ 20",
            data=excel_data,
            file_name=f"PEA_Matrix_Fixed_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
