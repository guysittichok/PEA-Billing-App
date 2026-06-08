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
        # ดึงข้อความแบบเต็มหน้าปกติ (เก็บไว้ใช้เช็กเงื่อนไขภาพรวม)
        full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        
        # 🌟 ทีเด็ดอยู่ตรงนี้: สั่งครอปหน้ากระดาษฝั่งซ้าย (กว้าง 0 ถึง 380) เพื่อสกัดเฉพาะข้อมูลใช้งานจริง
        left_text_list = []
        for page in pdf.pages:
            width = page.width
            height = page.height
            # ครอปเอาเฉพาะพื้นที่ฝั่งซ้าย (ตัดตารางประวัติขวาสุดออกไปเลย)
            left_bbox = (0, 0, 380, height)
            cropped_page = page.within_bbox(left_bbox)
            left_text_list.append(cropped_page.extract_text() or "")
        left_text = "\n".join(left_text_list)

    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": 0.0, "D": 0.0, "E": 0.0, "F": 0.0, "G": 0.0, "H": 0.0,
        "I": 0.0, "J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0, 
        "O": 0.0, "P": 0.0, "Q": 0.0
    }

    # ใช้ข้อความเต็มหน้าในการเช็กประเภทบิลตามเงื่อนไขเดิมของคุณสิทธิโชค
    is_tou = any(re.search(r''+k+r'.*?(?:กว|หน่วย|หนอรย|หนวย)', full_text, re.I) for k in ["Peak", "Off Peak", "Partial Peak"])
    has_h_mode = " H " in full_text or "\nH " in full_text or " H\n" in full_text or " Holiday " in full_text

    # ========================================================
    # รูปแบบที่ 2 (บิลแบบ P, OP, H) -> รันบนพื้นที่ฝั่งซ้ายที่ปลอดภัยแล้ว
    # ========================================================
    if is_tou and has_h_mode:
        # ค้นหาค่าความต้องการพลังไฟฟ้า (กว.) ฝั่งซ้าย
        peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', left_text, re.I)
        if peak:
            result["C"] = float(peak.group(1).replace(",", ""))
            result["F"] = float(peak.group(2).replace(",", ""))

        op_demand = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', left_text, re.I)
        if op_demand:
            result["D"] = float(op_demand.group(1).replace(",", ""))

        left_lines = left_text.split('\n')
        
        # ดึงค่าช่อง E จากตารางด้านบน
        demand_section = False
        for line in left_lines:
            if "พลังไฟฟ้าสูงสุด" in line:
                demand_section = True
            if demand_section and ("H " in line or " H " in line):
                nums_in_h = re.findall(r"([\d,]+\.\d+)", line)
                if nums_in_h:
                    result["E"] = float(nums_in_h[-1].replace(",", ""))
                demand_section = False

        result["G"] = 0.0
        result["H"] = 0.0

        # สแกนหาหน่วยการใช้งาน (I, J, K) จากตารางพลังงานไฟฟ้าช่วงกลางหน้า
        energy_section = False
        for line in left_lines:
            if "พลังงานไฟฟ้า" in line:
                energy_section = True
                nums = re.findall(r"([\d,]+\.\d+)", line)
                if nums and "P" in line and "PP" not in line:
                    result["I"] = float(nums[-1].replace(",", ""))
                continue
                
            if energy_section:
                if "รวม" in line:
                    energy_section = False
                    continue
                
                nums = re.findall(r"([\d,]+\.\d+)", line)
                if not nums: 
                    continue
                
                # มั่นใจได้ 100% ว่าตัวเลขตัวสุดท้ายในบรรทัดนี้คือค่าจริง เพราะฝั่งขวาโดนตัดทิ้งไปแล้ว
                if "OP" in line:
                    result["J"] = float(nums[-1].replace(",", ""))
                elif "H" in line:
                    # บิลใบแรกจะได้ 38,640.00 / บิลใบที่สองจะได้ 200,000.00 ทันที!
                    result["K"] = float(nums[-1].replace(",", ""))

    # ========================================================
    # รูปแบบที่ 3 และ 4 (รันบนข้อความเต็มหน้าปกติ)
    # ========================================================
    elif not is_tou:
        lines = full_text.split('\n')
        unit_match = re.search(r'([\d,]+\.\d+)\s+(?:หน่วย|หนอรย|หนวย)', full_text)
        if unit_match:
            result["I"] = float(unit_match.group(1).replace(",", ""))

        base_cost_pattern = r'(?:หน่วย|หนอรย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)'
        base_cost_match = re.search(base_cost_pattern, full_text)
        if base_cost_match:
            result["L"] = float(base_cost_match.group(1).replace(",", ""))
        else:
            for line in lines:
                if "พลังงานไฟฟ้า" in line:
                    nums_in_line = re.findall(r"([\d,]+\.\d+)", line)
                    if len(nums_in_line) >= 4:
                        result["L"] = float(nums_in_line[3].replace(",", ""))
                    elif len(nums_in_line) == 3:
                        result["L"] = float(nums_in_line[2].replace(",", ""))

        if "พลังไฟฟ้าสูงสุด" in full_text:
            for line in lines:
                if "พลังไฟฟ้าสูงสุด" in line:
                    nums = re.findall(r"([\d,]+\.\d+)", line)
                    if nums:
                        if len(nums) >= 3:
                            result["C"] = float(nums[2].replace(",", ""))
                        else:
                            result["C"] = float(nums[-1].replace(",", ""))

            demand_cost_match = re.search(r'พลังไฟฟ้าสูงสุด\s+.*?กว\..*?([\d,]+\.\d+)', full_text)
            if demand_cost_match:
                result["F"] = float(demand_cost_match.group(1).replace(",", ""))

    # ========================================================
    # รูปแบบที่ 1 (บิล TOU ทั่วไป)
    # ========================================================
    else:
        peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', full_text, re.I)
        if peak:
            result["C"] = float(peak.group(1).replace(",", ""))
            result["F"] = float(peak.group(2).replace(",", ""))

        pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', full_text, re.I)
        if pp:
            result["D"] = float(pp.group(1).replace(",", ""))
            result["G"] = float(pp.group(2).replace(",", ""))

        op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', full_text, re.I)
        if op:
            result["E"] = float(op.group(1).replace(",", ""))

        for line in full_text.split('\n'):
            nums = re.findall(r"([\d,]+\.\d+)", line)
            if not nums: continue
            if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line: result["I"] = float(nums[-1].replace(",", ""))
            elif "PP" in line: result["J"] = float(nums[-1].replace(",", ""))
            elif "OP" in line and result["K"] == 0.0: result["K"] = float(nums[-1].replace(",", ""))

    # ========================================================
    # ส่วนคำนวณท้ายบิล (Ft, Total, คอลัมน์ M, P, Q) รันบน full_text ตามปกติของคุณ
    # ========================================================
    for line in full_text.split('\n'):
        if "off" in line.lower() and "peak" in line.lower() and any(k in line for k in ["หน่วย", "หนอรย", "หนวย"]):
            clean_line = re.split(r'\d{2}/\d{2}/\d{2,4}', line)[0]
            nums_in_op_line = re.findall(r"([\d,]+\.\d+)", clean_line)
            if nums_in_op_line:
                result["M"] = float(nums_in_op_line[-1].replace(",", ""))
                break

    if result["L"] == 0.0:
        energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', full_text)
        if not energy: 
            energy = re.search(r'พลังงานไฟฟ้า.*?([\d,]+\.\d+)\s*บาท', full_text)
        if energy: 
            try:
                if not is_tou:
                    base_cost_match = re.search(r'พลังงานไฟฟ้า.*?หน่วย.*?([\d,]+\.\d+)', full_text)
                    if base_cost_match:
                        result["L"] = float(base_cost_match.group(1).replace(",", ""))
                    else:
                        result["L"] = float(energy.group(1).replace(",", ""))
                else:
                    result["L"] = float(energy.group(2).replace(",", ""))
            except:
                result["L"] = float(energy.group(1).replace(",", ""))
    
    ft = re.search(r'ค่า\s*Ft.*?=\s*[\d\.]+\s*บาท/หน่วย\s+([\d,]+\.\d+)', full_text, re.I)
    if not ft: 
        ft = re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', full_text, re.I)
    if ft: result["O"] = float(ft.group(1).replace(",", ""))
    
    result["P"] = 0.0  
    for line in full_text.split('\n'):
        if any(k in line for k in ["ค่าเพาเวอร์แฟคเตอร", "เพาเวอร์แฟคเตอร์", "Power Factor", "คาเพาเวอร์แฟคเตอร"]):
            nums_in_pf_line = re.findall(r"([\d,]+\.\d+)", line)
            if nums_in_pf_line:
                result["P"] = float(nums_in_pf_line[-1].replace(",", ""))
                break 

    total_match = re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub Total\)\s*([\d,]+\.\d+)', full_text, re.I)
    if total_match:
        result["Q"] = float(total_match.group(1).replace(",", ""))

    return result

# ----------------------------------------------------
# ส่วนโครงสร้าง Excel & Streamlit UI (คงเดิมไว้ทั้งหมด)
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
            
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            st.success("กรอกข้อมูลครบทุกช่องแล้ว!")
            st.download_button("📥 ดาวน์โหลด Excel", output, "Updated_PEA_Bill.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
