import streamlit as st
import datetime

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="centered")

st.title("ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")


st.divider()

st.sidebar.header("เลือกงวดประจำเดือน")
months_th = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]
current_year = datetime.datetime.now().year

selected_month = st.sidebar.selectbox("เลือกเดือน:", months_th, index=4)
selected_year = st.sidebar.number_input("เลือกปี (ค.ศ.):", min_value=2020, max_value=2040, value=current_year)
selected_date = st.sidebar.date_input("วันที่ออกรายงาน:", datetime.date.today())

st.sidebar.success(f"งวดที่เลือก: {selected_month} {selected_year}")

st.subheader("อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader(
    f"ลากไฟล์บิล PDF ประจำงวด [{selected_month} {selected_year}] มาวางที่นี่", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"ได้รับไฟล์บิลทั้งหมด {len(uploaded_files)} ไฟล์เรียบร้อยแล้ว!")
