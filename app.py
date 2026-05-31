import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันล็อกพิกัดตามบรรทัดบิลจริง (Line-Context Filter) - แก้ไขปัญหาสลับช่องตัวเลข")

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
                
    # ตั้งค่าตัวแปรช่องผลลัพธ์ C ถึง Q ให้ว่างเปล่าทั้งหมดเป็นค่าเริ่มต้น
    col_c = col_d = col_e = col_f = col_g = col_h = col_i = col_j = col_k = col_l = col_m = col_n = col_o = col_p = col_q = ""
    
    for line in text_lines:
        line_clean = line.strip()
        # ค้นหาตัวเลขทั้งหมดในบรรทัดนั้นๆ ที่เป็นตัวเลขทศนิยม
        nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line_clean)
        if not nums:
            continue
            
        # 1. ตรวจสอบกลุ่มความต้องการพลังงานไฟฟ้าสูงสุด (Demand Charge) ด้านบนของบิล
        if "กว." in line_clean or "กว" in line_clean:
            if "Peak" in line_clean and "Partial" not in line_clean and "Off" not in line_clean:
                col_c = clean_num(nums[0])  # ค่า Peak kW -> 3,620.00
                if len(nums) >= 2:
                    # ถ้ามีมากกว่า 1 ตัว ตัวท้ายสุดในบรรทัดนี้คือจำนวนเงินค่า Demand
                    col_f = clean_num(nums[-1]) # เงิน Demand Peak -> 1,031,881.00
            elif "Partial" in line_clean:
                col_d = clean_num(nums[0])  # ค่า PP kW -> 5,260.00
                if len(nums) >= 2:
                    col_g = clean_num(nums[-1]) # เงิน Demand PP -> 96,563.20
            elif "Off" in line_clean:
                col_e = clean_num(nums[0])  # ค่า Off Peak kW -> 5,240.00

        # 2. ตรวจสอบกลุ่มปริมาณการใช้ไฟฟ้าและค่าพลังงาน (Energy Charge)
        else:
            if "Peak" in line_clean and "Partial" not in line_clean and "Off" not in line_clean:
                # ในบรรทัดหน่วยใช้ไฟฟ้าสะสม ตัวแรกสุดคือจำนวนหน่วยที่ใช้จริง
                col_i = clean_num(nums[0])  # หน่วย Peak kWh -> 144,000.00
            elif "Partial" in line_clean:
                col_j = clean_num(nums[0])  # หน่วย PP kWh -> 651,000.00
            elif "Off" in line_clean:
                col_k = clean_num(nums[0])  # หน่วย Off Peak kWh -> 440,200.00
                if len(nums) >= 2:
                    # สำหรับ Off Peak ตัวที่สองจะเป็นจำนวนเงินรวมของค่าพลังงานไฟฟ้า
                    col_l = clean_num(nums[1]) # เงินพลังงาน Off Peak -> 3,887,297.92

        # 3. ตรวจสอบกลุ่มค่า Power Factor (ท้ายบิล)
        if "คาเพาเวอร์" in line_clean or "เพาเวอร์แฟคเตอร์" in line_clean or "Factor" in line_clean:
            col_p = clean_num(nums[0])      # เงินค่า Power Factor -> 584,249.40

    return {
        "ชื่อไฟล์": file_obj.name,
        "C": col_c, "D": col_d, "E": col_e,
        "F": col_f, "G": col_g, "H": col_h,
        "I": col_i, "J": col_j, "K": col_k, "L": col_l,
        "M": col_m, "N": col_n, "O": col_o,
        "P": col_p, "Q": col_q
    }

# การจัดการหน้ารายงาน Streamlit Web UI
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังแมปพิกัดข้อมูลป้องกันการสลับช่อง... {f.name}"):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ คัดกรองตัวเลขลงตำแหน่งช่อง C - Q ถูกต้องตามโครงสร้างตารางหลักแล้ว!")
        st.subheader("📊 2. ตารางพรีวิวก่อน Copy (ข้อมูลจะเรียงแถวราบลงล็อกพอดี)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Line_Context_Map')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับ Copy แปะลง Column C แถวที่ 20",
            data=excel_data,
            file_name=f"PEA_Line_Context_Fixed_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
