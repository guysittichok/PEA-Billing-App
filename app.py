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

    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": 0.0, "D": 0.0, "E": 0.0, "F": 0.0, "G": 0.0, "H": 0.0,
        "I": 0.0, "J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0, 
        "O": 0.0, "P": 0.0, "Q": 0.0
    }

    lines = text.split('\n')

    # 🌟 เช็กประเภทบิลแบบแม่นยำสูง (บิลประเภทที่ 2 ที่มีปัญหาซ้ำซ้อน)
    is_tou_h_mode = "พลังงานไฟฟ้า P" in text or ("OP" in text and "ประวัติการใช้ไฟฟ้า" in text)

    # ========================================================
    # 🎯 บล็อกสำหรับบิลประเภทที่ 2 (บิล TOU ที่มี Holiday ของ ปตท.)
    # ========================================================
    if is_tou_h_mode:
        # 1. ค้นหา Demand ท่อนบน (C, D, E) และ เงินค่า Demand (F, G, H) จากพิกัดแถว
        for line in lines:
            nums = re.findall(r"([\d,]+\.\d+)", line)
            if not nums: continue
            
            # ช่อง C: Peak กว. (ตารางซ้ายบรรทัดแรกสุด)
            if "Peak" in line and "กว" in line and "Off" not in line:
                result["C"] = float(nums[0].replace(",", ""))
                if len(nums) >= 2:
                    result["F"] = float(nums[-1].replace(",", ""))
                    
            # ช่อง D: OP กว. (ตารางซ้ายบรรทัดสอง) -> ดึงตัวแรกสุด มั่นใจได้ค่า 426.00 ไม่หลุดไปฝั่งขวา
            elif "OP" in line and "กว" in line:
                result["D"] = float(nums[0].replace(",", ""))
                if len(nums) >= 2:
                    result["G"] = float(nums[-1].replace(",", ""))
                    
            # ช่อง E: H กว. (ตารางซ้ายบรรทัดสาม) -> ดึงเลขตัวที่ 3 ของบรรทัดสถิติใช้งานจริง
            elif (line.strip().startswith("H ") or " H " in line) and "กว" in line:
                if len(nums) >= 3:
                    result["E"] = float(nums[2].replace(",", ""))
                else:
                    result["E"] = float(nums[0].replace(",", ""))

        # 2. ค้นหาตารางการใช้งาน "หน่วย" ท่อนล่างฝั่งซ้าย (I, J, K) 
        in_energy_zone = False
        for line in lines:
            if "พลังงานไฟฟ้า" in line:
                in_energy_zone = True
                nums = re.findall(r"([\d,]+\.\d+)", line)
                # ดึงหน่วยใช้งานจริงบรรทัดแรก (Peak หน่วย -> ช่อง I)
                if nums and "P" in line and "PP" not in line:
                    if len(nums) >= 3: result["I"] = float(nums[2].replace(",", ""))
                    else: result["I"] = float(nums[-1].replace(",", ""))
                continue
            
            # ชนคำว่า "รวม" แปลว่าหมดเขตตารางซ้าย ให้ปิดโซนทันทีเพื่อไม่ให้เลขฝั่งขวาหลอก
            if in_energy_zone and "รวม" in line:
                in_energy_zone = False
                break
                
            if in_energy_zone:
                nums = re.findall(r"([\d,]+\.\d+)", line)
                if not nums: continue
                
                # ช่อง J: Off Peak หน่วย (ตารางซ้ายบรรทัดสอง)
                if "OP" in line:
                    if len(nums) >= 3: result["J"] = float(nums[2].replace(",", ""))
                    else: result["J"] = float(nums[-1].replace(",", ""))
                # ช่อง K: Holiday หน่วย (ตารางซ้ายบรรทัดสาม) -> ดึง 38,640.00 ได้ชัวร์
                elif "H" in line:
                    if len(nums) >= 3: result["K"] = float(nums[2].replace(",", ""))
                    else: result["K"] = float(nums[-1].replace(",", ""))

        # 3. เจาะจงหาค่าไฟฟ้าฐานรวม (ช่อง L) จากบรรทัดรวมด้านขวา
        for line in lines:
            if "รวมเงินค่าไฟฟ้าฐาน" in line:
                nums = re.findall(r"([\d,]+\.\d+)", line)
                if nums:
                    result["L"] = float(nums[-1].replace(",", ""))
                    break

    # ========================================================
    # บิลงวดรูปแบบปกติอื่นๆ (คงตรรกะเดิมที่พี่รันผ่านไว้ทั้งหมด)
    # ========================================================
    else:
        is_tou_old = any(re.search(r''+k+r'.*?(?:กว|หน่วย|หนอรย|หนวย)', text, re.I) for k in ["Peak", "Off Peak", "Partial Peak"])
        
        if not is_tou_old:
            unit_match = re.search(r'([\d,]+\.\d+)\s+(?:หน่วย|หนอรย|หนวย)', text)
            if unit_match: result["I"] = float(unit_match.group(1).replace(",", ""))

            base_cost_pattern = r'(?:หน่วย|หนอรย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)'
            base_cost_match = re.search(base_cost_pattern, text)
            if base_cost_match:
                result["L"] = float(base_cost_match.group(1).replace(",", ""))
            else:
                for line in lines:
                    if "พลังงานไฟฟ้า" in line:
                        nums_in_line = re.findall(r"([\d,]+\.\d+)", line)
                        if len(nums_in_line) >= 4: result["L"] = float(nums_in_line[3].replace(",", ""))
                        elif len(nums_in_line) == 3: result["L"] = float(nums_in_line[2].replace(",", ""))

            if "พลังไฟฟ้าสูงสุด" in text:
                for line in lines:
                    if "พลังไฟฟ้าสูงสุด" in line:
                        nums = re.findall(r"([\d,]+\.\d+)", line)
                        if nums:
                            if len(nums) >= 3: result["C"] = float(nums[2].replace(",", ""))
                            else: result["C"] = float(nums[-1].replace(",", ""))
                demand_cost_match = re.search(r'พลังไฟฟ้าสูงสุด\s+.*?กว\..*?([\d,]+\.\d+)', text)
                if demand_cost_match: result["F"] = float(demand_cost_match.group(1).replace(",", ""))
        else:
            peak = re.search(r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
            if peak:
                result["C"] = float(peak.group(1).replace(",", ""))
                result["F"] = float(peak.group(2).replace(",", ""))
            pp = re.search(r'Partial\s+Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
            if pp:
                result["D"] = float(pp.group(1).replace(",", ""))
                result["G"] = float(pp.group(2).replace(",", ""))
            op = re.search(r'Off\s+Peak\s+([\d,]+\.\d+)\s+กว', text, re.I)
            if op: result["E"] = float(op.group(1).replace(",", ""))

            for line in text.split('\n'):
                nums = re.findall(r"([\d,]+\.\d+)", line)
                if not nums: continue
                if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line: result["I"] = float(nums[-1].replace(",", ""))
                elif "PP" in line: result["J"] = float(nums[-1].replace(",", ""))
                elif "OP" in line: result["K"] = float(nums[-1].replace(",", ""))

    # ========================================================
    # ส่วนท้ายดึงข้อมูลร่วมกัน (M, O, P, Q) -> สกัดเฉพาะบรรทัดและใช้พิกัดท้ายแถว
    # ========================================================
    # ช่อง M (เงิน Off Peak) คงตรรกะเดิมที่เคยหาเจอสำเร็จไว้
    for line in lines:
        if "off" in line.lower() and "peak" in line.lower() and any(k in line for k in ["หน่วย", "หนอรย", "หนวย", "บาท"]):
            clean_line = re.split(r'\d{2}/\d{2}/\d{2,4}', line)[0]
            nums_in_op_line = re.findall(r"([\d,]+\.\d+)", clean_line)
            if nums_in_op_line:
                result["M"] = float(nums_in_op_line[-1].replace(",", ""))
                break

    # ดึงเงินค่า Ft (ช่อง O)
    ft = re.search(r'ค่า\s*Ft.*?=\s*[\d\.]+\s*บาท/หน่วย\s+([\d,]+\.\d+)', text, re.I)
    if not ft: ft = re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I)
    if ft: result["O"] = float(ft.group(1).replace(",", ""))
    
    # ดึงค่า Power Factor (ช่อง P)
    for line in lines:
        if any(k in line for k in ["ค่าเพาเวอร์แฟคเตอร", "เพาเวอร์แฟคเตอร์", "Power Factor", "คาเพาเวอร์แฟคเตอร"]):
            nums_in_pf_line = re.findall(r"([\d,]+\.\d+)", line)
            if nums_in_pf_line:
                result["P"] = float(nums_in_pf_line[-1].replace(",", ""))
                break 

    # ยอดรวมเงินค่าไฟฟ้า (ช่อง Q)
    total_match = re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub Total\)\s*([\d,]+\.\d+)', text, re.I)
    if total_match: result["Q"] = float(total_match.group(1).replace(",", ""))

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
                        if val == 0: ws[cell_pos] = "-"
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
