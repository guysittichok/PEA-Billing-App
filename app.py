import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชัน Regex-Strict Mapping: สกัดค่าตรงตามเงื่อนไขของบิล ปตท. (แก้ไขบั๊กหน้าจอขาว)")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
current_year = datetime.datetime.now().year
selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=3) # เมษายน
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        # ใช้ \n เชื่อมต่อข้อความตามที่คุณเขียน
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    def get_float(pattern, text_source):
        # ปรับการใส่ Flags re.S และ re.I ให้ถูกต้องตามหลัก Python เพื่อป้องกันเว็บค้าง
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

    # 1. -------------------------
    # กลุ่ม Demand Charge (ความต้องการพลังงานสูงสุด)
    # -------------------------
    # ค้นหา Peak kW (C) และ เงิน Demand Peak (F)
    peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.IGNORECASE)
    if peak:
        result["C"] = float(peak.group(1).replace(",", ""))
        result["F"] = float(peak.group(2).replace(",", ""))

    # ค้นหา Partial Peak kW (D) และ เงิน Demand PP (G)
    pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.IGNORECASE)
    if pp:
        result["D"] = float(pp.group(1).replace(",", ""))
        result["G"] = float(pp.group(2).replace(",", ""))

    # ค้นหา Off Peak kW (E)
    op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.IGNORECASE)
    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # 2. -------------------------
    # กลุ่ม จำนวนหน่วยไฟฟ้าสะสม (Usage Section) จากตารางสถิติด้านล่าง
    # -------------------------
    # ดักจับจากสัญลักษณ์ P, PP, OP คู่กับกลุ่มตัวเลขเพื่อความแม่นยำสูง
    usage_peak = re.search(r'P\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_peak:
        result["I"] = float(usage_peak.group(1).replace(",", ""))

    usage_pp = re.search(r'PP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_pp:
        result["J"] = float(usage_pp.group(1).replace(",", ""))

    usage_op = re.search(r'OP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if usage_op:
        result["K"] = float(usage_op.group(1).replace(",", ""))

    # 3. -------------------------
    # กลุ่มเงินค่าพลังงานไฟฟ้า และรวมหน่วย (Energy Cost)
    # -------------------------
    # ปรับให้รองรับคำว่า 'หนวย' หรือ 'หน่วย' ที่สระมักเพี้ยนใน PDF
    energy = re.search(r'([\d,]+\.\d+)\s+หนอรย|หน่วย|หนวย\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", "")) # รวมจำนวนหน่วย (O)
        result["L"] = float(energy.group(2).replace(",", "")) # เงินค่าพลังงานไฟฟ้าช่วง Off Peak / รวมท้าย (L)

    # 4. -------------------------
    # กลุ่มค่าใช้จ่ายอื่น ๆ ท้ายบิล (Ft, Service, Power Factor, Sub Total)
    # -------------------------
    # ค่า Ft
    result["M"] = get_float(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text)
    
    # ค่าบริการรายเดือน
    result["N"] = get_float(r'ค่าบริการ.*?([\d,]+\.\d+)', text)
    
    # เงินค่า Power Factor (ดักคำว่า คาเพาเวอร์แฟคเตอร แบบไม่มีวรรณยุกต์ตามบิล PEA จริง)
    result["P"] = get_float(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text)
    
    # เงินรวม Sub Total
    result["Q"] = get_float(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text)

    return result

# หน้าจอการทำงานหลักของเว็บแอป Streamlit
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสกัดข้อมูลระบบ Regex Strict บล็อกต่อบล็อก... {f.name}"):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ ประมวลผลเสร็จสิ้น! ไม่พบบั๊กหน้าจอขาว")
        st.subheader("📊 2. ตารางพรีวิวก่อนคัดลอก (ตรวจสอบพิกัดช่อง C ถึง Q)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='PEA_Strict_Report')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำหรับก๊อปปี้ไปแปะในหน้าหลัก",
            data=excel_data,
            file_name=f"PEA_Strict_Extracted_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
