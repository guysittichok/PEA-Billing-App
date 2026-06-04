import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
import openpyxl

st.set_page_config(page_title="ระบบจัดการบิลค่าไฟฟ้า", layout="wide")
st.title("PEA Bill Extraction System")

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    # ตั้งค่าเริ่มต้นเป็น 0.0 ตามแบบเดิมของคุณ
    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": 0.0, "D": 0.0, "E": 0.0, "F": 0.0, "G": 0.0, "H": 0.0,
        "I": 0.0, "J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0, 
        "O": 0.0, "P": 0.0, "Q": 0.0
    }

    # ----------------------------------------------------
    # ส่วนที่ 1: ตรวจสอบและแบ่งกลุ่มรูปแบบบิล (Bill Type Identification)
    # ----------------------------------------------------
    is_tou = any(x in text for x in ["Peak", "Off Peak", "Partial Peak", "PP", "OP"])
    has_holiday = "H " in text or "\nH" in text or " H" in text

    # ----------------------------------------------------
    # ส่วนที่ 2: ดึงข้อมูลตามรูปแบบบิลที่ตรวจสอบได้
    # ----------------------------------------------------
    
    # ==== รูปแบบที่ 1: บิล TOU ปกติ (มี P, PP, OP) ====
    if is_tou and not has_holiday:
        # 1. Demand Charge (C, D, E, F, G)
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

        # 2. Energy Usage (I, J, K)
        for line in text.split('\n'):
            nums = re.findall(r"([\d,]+\.\d+)", line)
            if not nums: continue
            if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line: result["I"] = float(nums[-1].replace(",", ""))
            elif "PP" in line: result["J"] = float(nums[-1].replace(",", ""))
            elif "OP" in line: result["K"] = float(nums[-1].replace(",", ""))

    # ==== รูปแบบที่ 2: บิล TOU แบบพิเศษ (มี P, OP, H) ====
    elif is_tou and has_holiday:
        # 1. Demand Charge (จัดเรียงตามลำดับ แต่อดกรอกช่อง G)
        peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
        if peak:
            result["C"] = float(peak.group(1).replace(",", ""))
            result["F"] = float(peak.group(2).replace(",", ""))

        # แถวที่ 2 เป็น Off Peak ยิงเข้าช่อง D (ส่วนช่อง G ปล่อยเป็น 0.0 ไม่กรอกค่า)
        op_demand = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
        if op_demand:
            result["D"] = float(op_demand.group(1).replace(",", ""))

        # แถวที่ 3 เป็น H (Holiday) ยิงเข้าช่อง E และ H
        h_demand = re.search(r'(?:\s+|\n|^)H\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
        if h_demand:
            result["E"] = float(h_demand.group(1).replace(",", ""))
            result["H"] = float(h_demand.group(2).replace(",", ""))

        # 2. Energy Usage (I, J, K เรียงตามลำดับลงมา)
        for line in text.split('\n'):
            nums = re.findall(r"([\d,]+\.\d+)", line)
            if not nums: continue
            if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line: result["I"] = float(nums[-1].replace(",", ""))
            elif "OP" in line: result["J"] = float(nums[-1].replace(",", ""))
            elif "H " in line or line.strip().startswith("H ") or " H " in line: result["K"] = float(nums[-1].replace(",", ""))

    # ==== รูปแบบที่ 3 และ 4: บิลอัตราปกติ (ภาษาไทยล้วน ไม่มี P, PP, OP, H) ====
    else:
        # ค้นหาว่าในบิลมีแถว "พลังไฟฟ้าสูงสุด" หรือไม่
        has_max_demand = "พลังไฟฟ้าสูงสุด" in text
        
        if has_max_demand:
            # ==== รูปแบบที่ 3: อัตราปกติแบบมี พลังไฟฟ้าสูงสุด (กรอก C และ I) ====
            demand_match = re.search(r'พลังไฟฟ้าสูงสุด\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
            if demand_match:
                result["C"] = float(demand_match.group(1).replace(",", ""))
                result["F"] = float(demand_match.group(2).replace(",", ""))
                
            energy_match = re.search(r'พลังงานไฟฟ้า\s+([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)', text)
            if energy_match:
                result["I"] = float(energy_match.group(1).replace(",", ""))
        else:
            # ==== รูปแบบที่ 4: อัตราปกติแบบไม่มี พลังไฟฟ้าสูงสุด (มีบรรทัดเดียว กรอกลงช่อง I) ====
            energy_match = re.search(r'พลังงานไฟฟ้า\s+([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)', text)
            if energy_match:
                result["I"] = float(energy_match.group(1).replace(",", ""))

    # ----------------------------------------------------
    # ส่วนที่ 3: ดึงค่าใช้จ่าย และ ค่าอื่นๆ ด้านล่างบิล (ใช้ได้กับทุกรูปแบบ)
    # ----------------------------------------------------
    energy_cost = re.search(r'(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy_cost: result["L"] = float(energy_cost.group(1).replace(",", ""))
    
    ft = re.search(r'ค่า\s*Ft.*?=\s*[-\d\.]+\s*บาท/หน่วย\s+([\d,]+\.\d+)', text, re.I)
    if not ft:  # ดักสำรองกรณีรูปแบบข้อความ Ft แตกต่างออกไป
        ft = re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I)
    if ft: result["O"] = float(ft.group(1).replace(",", ""))
    
    pf = re.search(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text, re.I)
    if pf: result["P"] = float(pf.group(1).replace(",", ""))
    
    total_match = re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub Total\)\s*([\d,]+\.\d+)', text, re.I)
    if total_match:
        result["Q"] = float(total_match.group(1).replace(",", ""))

    return result

# ----------------------------------------------------
# ส่วนที่ 4: Streamlit UI & Excel Mapping (โครงสร้างเดิมของคุณ)
# ----------------------------------------------------
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)
template_file = st.file_uploader("อัปโหลดไฟล์ Excel ต้นแบบ", type=["xlsx"])

if uploaded_files:
    data = [extract_exact_pea_bill(f) for f in uploaded_files]
    all_cols = ["ชื่อไฟล์", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    df = pd.DataFrame(data, columns=all_cols)
    st.data_editor(df, use_container_width=True)

    if template_file and st.button("สร้างไฟล์ Excel พร้อมข้อมูล"):
        try:
            wb = openpyxl.load_workbook(template_file)
            ws = wb.active 

            # 1. ฟังก์ชัน write_number (กำหนดไว้ในนี้ตามแบบของคุณ)
            def write_number(ws, cell_pos, value):
                val_str = str(value).strip()
                if val_str in ["0", "0.0", "None", ""]:
                    ws[cell_pos] = "-"
                else:
                    try:
                        val = float(val_str.replace(',', ''))
                        if val == 0:
                            ws[cell_pos] = "-"
                        else:
                            ws[cell_pos] = val
                            ws[cell_pos].number_format = '0.00'
                    except:
                        ws[cell_pos] = "-"

            # 2. ส่วนการวนลูปกรอกข้อมูลลงแถวที่ 20 เป็นต้นไป
            for row_idx in range(20, ws.max_row + 1):
                excel_key = str(ws[f'A{row_idx}'].value).strip()
                if excel_key in ["None", ""]: continue
                
                match_row = df[df['ชื่อไฟล์'].apply(lambda x: excel_key in str(x))]
                
                if not match_row.empty:
                    row = match_row.iloc[0]
                    write_number(ws, f'C{row_idx}', row['C'])
                    write_number(ws, f'D{row_idx}', row['D'])
                    write_number(ws, f'E{row_idx}', row['E'])
                    write_number(ws, f'F{row_idx}', row['F'])
                    write_number(ws, f'G{row_idx}', row['G'])
                    write_number(ws, f'H{row_idx}', row['H'])
                    write_number(ws, f'I{row_idx}', row['I'])
                    write_number(ws, f'J{row_idx}', row['J'])
                    write_number(ws, f'K{row_idx}', row['K'])
                    write_number(ws, f'L{row_idx}', row['L'])
                    write_number(ws, f'M{row_idx}', row['M'])
                    write_number(ws, f'N{row_idx}', row['N'])
                    write_number(ws, f'O{row_idx}', row['O'])
                    write_number(ws, f'P{row_idx}', row['P'])
                    write_number(ws, f'Q{row_idx}', row['Q'])
            
            # 3. เซฟและสร้างปุ่มดาวน์โหลด
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.success("กรอกข้อมูลครบทุกช่องแล้ว!")
            st.download_button("📥 ดาวน์โหลด Excel", output, "Updated_PEA_Bill.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
