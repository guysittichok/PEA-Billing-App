import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันทำงานร่วมกับบิลจริง: ดึงข้อมูลด้วยดัชนีตัวเลขระดับโครงสร้าง (ไม่เจอกลุ่มตัวเลขเป็น 0)")

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
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        text_lines = text.split('\n')
                
    # ตั้งค่าตัวแปรผลลัพธ์ C ถึง Q ให้เป็นค่าว่างรอรับข้อมูลจริง
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    for line in text_lines:
        line_clean = line.strip()
        # ค้นหาตัวเลขทั้งหมดที่มีการใส่เครื่องหมายจุลภาคหรือทศนิยม เช่น 3,620.00 หรือ 144,000.00
        all_numbers = re.findall(r"([0-9,]+\.[0-9]{2,4})", line_clean)
        
        if not all_numbers:
            continue
            
        # 1. จัดการกลุ่ม Peak
        if "Peak" in line_clean and "Partial" not in line_clean and "Off" not in line_clean:
            # เช็กว่าเป็นบรรทัดคิดเงินด้านบน (มีอัตรา 285.0500 อยู่ด้วย)
            if "285.05" in line_clean or len(all_numbers) >= 3:
                col_c = clean_num(all_numbers[0])  # ค่า Peak kW -> 3,620.00 (คอลัมน์ C)
                col_f = clean_num(all_numbers[-1]) # เงินรวม Peak -> 1,031,881.00 (คอลัมน์ F)
            else:
                # ถ้าไม่มีอัตราแปลว่าเป็นบรรทัดหน่วยด้านล่าง
                col_i = clean_num(all_numbers[0])  # หน่วย Peak -> 144,000.00 (คอลัมน์ I)

        # 2. จัดการกลุ่ม Partial Peak
        elif "Partial" in line_clean:
            if "58.88" in line_clean or len(all_numbers) >= 3:
                col_d = clean_num(all_numbers[0])  # ค่า PP kW -> 5,260.00 (คอลัมน์ D)
                col_g = clean_num(all_numbers[-1]) # เงินรวม PP -> 96,563.20 (คอลัมน์ G)
            else:
                col_j = clean_num(all_numbers[0])  # หน่วย PP -> 651,000.00 (คอลัมน์ J)

        # 3. จัดการกลุ่ม Off Peak
        elif "Off" in line_clean:
            # เช็กว่าเป็นบรรทัดความต้องการพลังงานด้านบน (มีหน่วย กว. หรือตัวเลขชุดเดียว)
            if "กว" in line_clean or len(all_numbers) == 1:
                col_e = clean_num(all_numbers[0])  # ค่า Off Peak kW -> 5,240.00 (คอลัมน์ E)
            else:
                # เป็นบรรทัดตารางหน่วยด้านล่างที่มีตัวเลขเรียงกัน 2 ชุด
                if len(all_numbers) >= 2:
                    col_k = clean_num(all_numbers[0])  # หน่วย Off Peak -> 440,200.00 (คอลัมน์ K)
                    col_l = clean_num(all_numbers[1])  # เงินพลังงาน Off Peak -> 3,887,297.92 (คอลัมน์ L)

        # 4. จัดการเงินค่า Power Factor ท้ายบิล
        elif "เพาเวอร์" in line_clean or "แฟคเตอร์" in line_clean or "Factor" in line_clean:
            col_p = clean_num(all_numbers[0])      # เงินค่า Power Factor -> 584,249.40 (คอลัมน์ P)

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
        with st.spinner(f"กำลังสกัดข้อมูลเข้าตารางหลัก {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ ประมวลผลและกระจายตัวเลขลงคอลัมน์สำเร็จ!")
        st.subheader("📊 2. ตารางพรีวิวก่อนคัดลอก (เรียงช่อง C ถึง Q ตามแบบโครงสร้างบัญชี)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Final_Ready')
            return output.
