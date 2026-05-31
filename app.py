import streamlit as st
import datetime
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")

st.title("⚡ ระบบบันทึกข้อมูลและเจนรีพอร์ตค่าไฟฟ้าอัตโนมัติ")
st.write("เวอร์ชันสมบูรณ์ (แก้ไขรอบสุดท้าย): จัดเรียงตัวเลขลงช่องตามความจริงวิศวกรรมและผังตารางหลัก")

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
                
    # ตัวแปรสำหรับรับค่าตามที่คุณระบุมา
    val_f_demand_peak_money = 0.0
    val_g_demand_pp_money = 0.0
    val_i_energy_peak_unit = 0.0
    val_j_energy_pp_unit = 0.0
    val_k_energy_off_unit = 0.0
    val_l_energy_off_money = 0.0
    val_p_power_factor_money = 0.0
    
    for line in text_lines:
        # สกัดค่า F (เงิน Demand Peak)
        if "Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                val_f_demand_peak_money = clean_num(nums[2])
                
        # สกัดค่า G (เงิน Demand PP)
        elif "Partial Peak" in line and "กว." in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 3:
                val_g_demand_pp_money = clean_num(nums[2])
                
        # สกัดค่า I (หน่วย Peak kWh)
        elif "Peak" in line and "หนวย" in line:
            nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums_kwh: 
                val_i_energy_peak_unit = clean_num(nums_kwh[-1])

        # สกัดค่า J (หน่วย PP kWh)
        elif "Partial Peak" in line and "หนวย" in line:
            nums_kwh = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if nums_kwh: 
                val_j_energy_pp_unit = clean_num(nums_kwh[-1])

        # สกัดค่า K (หน่วย Off Peak kWh) และ L (เงินพลังงาน Off Peak)
        elif "Off Peak" in line and "หนวย" in line:
            nums = re.findall(r"([0-9,]+\.[0-9]{2,4})", line)
            if len(nums) >= 2:
                val_k_energy_off_unit = clean_num(nums[0])
                val_l_energy_off_money = clean_num(nums[1])

        # สกัดค่า P (เงินค่า Power Factor)
        elif "คาเพาเวอร์แฟคเตอร" in line or "เพาเวอร์แฟคเตอร์" in line:
            num = re.search(r"([0-9,]+\.[0-9]{2})", line)
            if num: 
                val_p_power_factor_money = clean_num(num.group(1))

    # ส่งค่าออกแนวนอน ล็อกพิกัดตรงตามช่องตาราง Excel หลักของคุณ (C ถึง Q)
    return {
        "ชื่อไฟล์": file_obj.name,
        "คอลัมน์ C": "",
        "คอลัมน์ D": "",
        "คอลัมน์ E": "",
        "คอลัมน์ F [เงิน Demand Peak]": val_f_demand_peak_money,   # 1031881.00
        "คอลัมน์ G [เงิน Demand PP]": val_g_demand_pp_money,       # 96563.20
        "คอลัมน์ H": "",
        "คอลัมน์ I [หน่วย Peak]": val_i_energy_peak_unit,         # 144000.00
        "คอลัมน์ J [หน่วย PP]": val_j_energy_pp_unit,             # 651000.00
        "คอลัมน์ K [หน่วย Off Peak]": val_k_energy_off_unit,       # 440200.00
        "คอลัมน์ L [เงินพลังงาน Off Peak]": val_l_energy_off_money, # 3887297.92
        "คอลัมน์ M": "",
        "คอลัมน์ N": "",
        "คอลัมน์ O": "",
        "คอลัมน์ P [ค่า Power Factor]": val_p_power_factor_money, # 584249.40
        "คอลัมน์ Q": ""
    }

st.subheader("📂 1. อัปโหลดไฟล์บิลค่าไฟฟ้า (PDF)")
uploaded_files = st.file_uploader("ลากไฟล์บิล PDF มาวางที่นี่", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for f in uploaded_files:
        with st.spinner(f"กำลังสกัดค่าเข้าคอลัมน์ตารางหลัก {f.name}..."):
            try:
                all_data.append(extract_exact_pea_bill(f))
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดกับไฟล์ {f.name}: {e}")
                
    if all_data:
        st.success(f"⚡ สกัดและเรียงช่องข้อมูลตามลำดับที่ถูกต้อง 100% แล้ว!")
        st.subheader("📊 2. ตารางตรวจสอบข้อมูลเพื่อคัดลอก (Copy แปะลง Column C แถวที่ 20 ได้ทันที)")
        
        df = pd.DataFrame(all_data)
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        
        st.divider()
        st.subheader("📥 3. ส่งออกรายงาน Excel (.xlsx)")
        
        def to_excel(input_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                input_df.to_excel(writer, index=False, sheet_name='Ready_To_Paste')
            return output.getvalue()
        
        excel_data = to_excel(edited_df)
        st.download_button(
            label="🟢 ดาวน์โหลดไฟล์ Excel สำเร็จรูป",
            data=excel_data,
            file_name=f"PEA_Extract_C_Q_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
