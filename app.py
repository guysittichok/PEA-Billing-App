import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชัน Regex-Strict สมบูรณ์: แก้ไขบั๊กเครื่องหมายกลุ่มคำ (OR) ที่ทำให้หน้าจอเว็บบอร์ดขาว")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    def get_float(pattern, text_source):
        m = re.search(pattern, text_source, re.DOTALL | re.IGNORECASE)
        if m:
            return float(m.group(1).replace(",", ""))
        return ""

    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": "", "D": "", "E": "", "F": "", "G": "", "H": "", 
        "I": "", "J": "", "K": "", "L": "", "M": "", "N": "", 
        "O": "", "P": "", "Q": ""
    }

    # -------------------------
    # Demand Charge (กลุ่มความต้องการพลังงานสูงสุด)
    # -------------------------
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if peak:
        result["C"] = float(peak.group(1).replace(",", ""))
        result["F"] = float(peak.group(2).replace(",", ""))

    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
    if pp:
        result["D"] = float(pp.group(1).replace(",", ""))
        result["G"] = float(pp.group(2).replace(",", ""))

    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # -------------------------
    # Service Charge (ค่าบริการรายเดือน)
    # -------------------------
    result["N"] = get_float(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text)

    # -------------------------
    # Total Unit + Energy Cost (มัดรวมกลุ่มคำว่าหน่วยไม่ให้แตกแถว)
    # -------------------------
    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", ""))
        result["L"] = float(energy.group(2).replace(",", ""))

    # -------------------------
    # Usage Section (ปริมาณการใช้แยกช่วงเวลาจากตารางสถิติ P, PP, OP)
    # -------------------------
    usage_peak = re.search(r'P\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_peak:
        result["I"] = float(usage_peak.group(1).replace(",", ""))

    usage_pp = re.search(r'PP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_pp:
        result["J"] = float(usage_pp.group(1).replace(",", ""))

    usage_op = re.search(r'OP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_op:
        result["K"] = float(usage_op.group(1).replace(",", ""))

    # -------------------------
    # FT (ค่า Ft ประจำงวด)
    # -------------------------
    result["M"] = get_float(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text)

    # -------------------------
    # Power Factor (เงินค่าตัวประกอบกำลัง)
    # -------------------------
    result["P"] = get_float(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text)

    # -------------------------
    # Sub Total (รวมเงินค่าไฟฟ้าก่อนภาษี)
    # -------------------------
    result["Q"] = get_float(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text)

    return result

# โครงสร้างการแสดงผลหน้าเว็บแอป (Streamlit UI)
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสกัดค่าเข้าสู่ล็อกพิกัดตารางหลัก... {f.name}"):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดภายในระบบดึงข้อมูลของไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ แก้ไขข้อผิดพลาดของกลุ่มโค้ดและประมวลผลสำเร็จ!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูลพรีวิว (คอลลัมน์ C ถึง Q เรียงราบพอดีเป๊ะ)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Fixed_Data')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์รายงานสำเร็จรูป",
            data=excel_data,
            file_name=f"PEA_Verified_Report_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
