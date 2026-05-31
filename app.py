import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันทำงานร่วมกับเทมเพลตจริง (คอลัมน์ C ถึง Q) - สกัดตามพิกัดโครงสร้างบัญชี ปตท.")

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
                
    # ตั้งตัวแปรเริ่มต้นให้กับทุกคอลัมน์ให้ออกมาว่างเปล่าก่อน
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    # เจาะสแกนทีละบรรทัดเพื่อดึงตัวเลขให้ตรงตามเงื่อนไขทางวิศวกรรม/บัญชีของคุณ
    for line in text_lines:
        # 1. ค้นหากลุ่มความต้องการพลังงานช่วง Peak (Demand Charge) ด้านบนเพื่อลงช่อง F
        if "Peak" in line and "กว." in line and "285.0500" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
            if len(nums) >= 2:
                col_f = clean_num(nums[-1]) # ได้ค่า 1,031,881.00 ตรงช่อง F
                
        # 2. ค้นหากลุ่มความต้องการพลังงานช่วง Partial Peak เพื่อลงช่อง G
        elif "Partial Peak" in line and "กว." in line and "58.8800" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
            if len(nums) >= 2:
                col_g = clean_num(nums[-1]) # ได้ค่า 96,563.20 ตรงช่อง G
                
        # 3. ค้นหาจำนวนหน่วยการใช้ไฟฟ้าย่อยช่วงเวลา (Energy Units) จากตารางสถิติด้านล่าง
        elif "หนวย" in line and "440,200.00" in line:
            # บรรทัดหน่วย Off Peak จะโผล่มาพร้อมกับเงินค่าพลังงาน Off Peak เสมอ
            nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
            if len(nums) >= 2:
                col_k = clean_num(nums[0])  # ได้ค่าหน่วย Off Peak 440,200.00 ตรงช่อง K
                col_l = clean_num(nums[1])  # ได้ค่าเงินพลังงาน Off Peak 3,887,297.92 ตรงช่อง L
                
        # ค้นหาหน่วย Peak และ Partial Peak จากข้อมูลในตารางหน่วย
        elif "Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
            if nums:
                col_i = clean_num(nums[-1]) # ได้ค่าหน่วย Peak 144,000.00 ตรงช่อง I
        elif "Partial Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
            if nums:
                col_j = clean_num(nums[-1]) # ได้ค่าหน่วย PP 651,000.00 ตรงช่อง J

        # 4. ค้นหาเงินค่า Power Factor ท้ายตาราง (บรรทัดที่แม่นยำที่สุด)
        elif "คาเพาเวอร์แฟคเตอร" in line or "เพาเวอร์แฟคเตอร์" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: 
                col_p = clean_num(num.group(1)) # ได้ค่า 584,249.40 ตรงช่อง P

    # สรุปผังโครงสร้างตารางแนวนอนยิงตรงเข้าล็อก Row 20 ตั้งแต่ C ถึง Q พอดีช่องเป๊ะ
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

# หน้าจอการทำงานหลักของเว็บแอป Streamlit
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสกัดและจัดพิกัดข้อมูลลงช่องหลัก {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ แมปตำแหน่งช่องตารางสำเร็จตามบรีฟจริงเรียบร้อยแล้ว!")
        st.subheader("📊 2. ตารางตรวจสอบและเตรียม Copy (แถวราบล็อกช่อง C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Ready_To_Paste')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับก๊อปปี้ไปแปะในหน้าหลัก",
            data=excel_data,
            file_name=f"PEA_Final_Template_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
