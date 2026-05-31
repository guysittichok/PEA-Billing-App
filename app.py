import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide") # ปรับหน้าจอให้กว้างขึ้นเพื่อใส่ตาราง

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("ยินดีต้อนรับ! ระบบสามารถอ่านไฟล์บิล PDF และสรุปข้อมูลออกมาได้อัตโนมัติ")

st.divider()

# แถบเมนูด้านซ้าย (Sidebar)
st.sidebar.header("📅 เลือกงวดประจำเดือน")
months_th = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
current_year = datetime.datetime.now().year

selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=4)
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)
selected_date = st.sidebar.date_input("วันที่ออกรายงาน:", datetime.date.today())

st.sidebar.success(f"งวดที่เลือก: {selected_month} {selected_year}")

# ฟังก์ชันหลังบ้านสำหรับแกะข้อความจาก PDF
def extract_bill_data(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        # รวมข้อความจากทุกหน้าในไฟล์ PDF นั้นๆ
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    # ดึงข้อมูลด้วย Regex จากคีย์เวิร์ดที่ให้มา
    ca_match = re.search(r"CA/Ref\.No\.1\s*\n*\s*([0-9]+)", text)
    inv_match = re.search(r"Invoice no\.\s*\n*\s*([0-9]+)", text)
    amount_match = re.search(r"Total\s*\(Baht\)\s*\n*\s*([0-9,]+\.[0-9]{2})", text)
    due_match = re.search(r"Due Date\s*\n*\s*([^\n]+)", text)
    
    return {
        "ชื่อไฟล์": file_obj.name,
        "หมายเลขผู้ใช้ไฟฟ้า (CA)": ca_match.group(1) if ca_match else "ไม่พบข้อมูล",
        "เลขที่ใบแจ้งค่าไฟฟ้า": inv_match.group(1) if inv_match else "ไม่พบข้อมูล",
        "จำนวนเงินรวม (บาท)": amount_match.group(1) if amount_match else "ไม่พบข้อมูล",
        "วันครบกำหนด": due_match.group(1).strip() if due_match else "ไม่พบข้อมูล"
    }

# หน้าจอหลัก: ช่องอัปโหลดไฟล์
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader(
    f"ลากไฟล์บิล PDF ประจำงวด [{selected_month} {selected_year}] มาวางที่นี่", 
    type=["pdf"], 
    accept_multiple_files=True
)

# ถ้ามีการอัปโหลดไฟล์เข้ามา
if uploaded_files:
    all_data = []
    
    # วนลูปอ่านทีละไฟล์ที่ผู้ใช้วางลงไป
    for f in uploaded_files:
        with st.spinner(f"กำลังอ่านข้อมูลจากไฟล์ {f.name}..."):
            try:
                data = extract_bill_data(f)
                all_data.append(data)
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ อ่านข้อมูลบิลสำเร็จทั้งหมด {len(all_data)} ไฟล์!")
        
        st.subheader("📊 2. ตารางตรวจสอบและตรวจทานข้อมูล")
        st.write("คุณสามารถดับเบิ้ลคลิกในช่องตารางด้านล่างนี้ เพื่อแก้ไขตัวเลขให้ถูกต้องก่อนบันทึกได้ครับ")
        
        # แปลงข้อมูลเป็นตาราง DataFrame
        df = pd.DataFrame(all_data)
        
        # แสดงผลเป็น Data Editor ที่สามารถกดแก้ไขบนหน้าจอเว็บได้โดยตรง
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        # เก็บข้อมูลที่ตรวจทานแล้วไว้ในระบบจำลอง
        st.session_state["processed_data"] = edited_df
