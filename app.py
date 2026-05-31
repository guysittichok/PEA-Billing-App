import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

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

# ฟังก์ชันหลังบ้านสแกนข้อความแบบยืดหยุ่นสูง
def extract_bill_data(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    
    # 1. ค้นหาหมายเลขผู้ใช้ไฟฟ้า (CA): มองหาตัวเลข 12 หลัก (มักขึ้นต้นด้วย 02)
    # หรือลองตัดเลข 12 หลักมาจากชื่อไฟล์ก่อนถ้าหาในเนื้อหาไม่เจอ
    ca_from_filename = re.search(r"0200[0-9]{8}", file_obj.name)
    ca_match = re.search(r"\b(0200[0-9]{8})\b", text)
    
    if ca_match:
        ca_number = ca_match.group(1)
    elif ca_from_filename:
        ca_number = ca_from_filename.group(0)
    else:
        ca_number = "ไม่พบข้อมูล"

    # 2. ค้นหาเลขที่ใบแจ้งค่าไฟฟ้า (Invoice No.): มักเป็นเลข 12 หลักที่อยู่ใกล้ๆ กัน
    inv_matches = re.findall(r"\b([0-9]{12})\b", text)
    invoice_no = "ไม่พบข้อมูล"
    if inv_matches:
        # ถ้ามีเลข 12 หลักหลายตัว ตัวที่ไม่ใช่ CA ก็คือ Invoice No.
        for num in inv_matches:
            if num != ca_number:
                invoice_no = num
                break
        if invoice_no == "ไม่พบข้อมูล":
            invoice_no = inv_matches[0]

    # 3. ค้นหาจำนวนเงินรวม: สแกนหาตัวเลขที่มีคอมมาคั่นหลักล้าน/หลักแสน และลงท้ายด้วยทศนิยม 2 ตำแหน่ง
    amount_matches = re.findall(r"\b([0-9,]+\.[0-9]{2})\b", text)
    total_amount = "ไม่พบข้อมูล"
    if amount_matches:
        # โดยปกติยอดเงินรวมสุทธิมักจะเป็นตัวเลขที่มีค่ามากที่สุดในกลุ่มเงิน หรืออยู่ท้ายๆ
        # ในที่นี้เราดึงตัวที่ตรงกับแพทเทิร์นเงินมาตรวจสอบ
        for amt in amount_matches:
            clean_amt = amt.replace(",", "")
            if float(clean_amt) > 1000: # กรองตัวเลขค่าน้อยๆ ออกไป
                total_amount = amt

    # 4. ค้นหาวันครบกำหนดชำระ (Due Date)
    due_match = re.search(r"(\d{1,2}\s+[ก-์]{3,10}\s+\d{4})", text)
    due_date = due_match.group(1) if due_match else "ไม่พบข้อมูล"
    
    return {
        "ชื่อไฟล์": file_obj.name,
        "หมายเลขผู้ใช้ไฟฟ้า (CA)": ca_number,
        "เลขที่ใบแจ้งค่าไฟฟ้า": invoice_no,
        "จำนวนเงินรวม (บาท)": total_amount,
        "วันครบกำหนด": due_date
    }

# ช่องอัปโหลดไฟล์
st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader(
    f"ลากไฟล์บิล PDF ประจำงวด [{selected_month} {selected_year}] มาวางที่นี่", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสแกนวิเคราะห์ไฟล์ {f.name}..."):
            try:
                data = extract_bill_data(f)
                all_data.append(data)
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สแกนและจัดกลุ่มข้อมูลสำเร็จ {len(all_data)} ไฟล์!")
        st.subheader("📊 2. ตารางตรวจสอบและตรวจทานข้อมูล")
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        st.session_state["processed_data"] = edited_df
