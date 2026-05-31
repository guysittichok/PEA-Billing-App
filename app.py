import streamlit as st
import pdfplumber
import pandas as pd
import openpyxl
from io import BytesIO

def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        all_tables = []
        for page in pdf.pages:
            all_tables.extend(page.extract_tables())

    # กำหนดค่าเริ่มต้นเป็น 0.0 ทุกช่อง
    result = {
        "ชื่อไฟล์": file_obj.name,
        "C": 0.0, "D": 0.0, "E": 0.0, "F": 0.0, "G": 0.0, "H": 0.0, 
        "I": 0.0, "J": 0.0, "K": 0.0, "L": 0.0, "M": 0.0, "N": 0.0, 
        "O": 0.0, "P": 0.0, "Q": 0.0
    }

    # 1. ดึงยอดเงินจาก Text (ใช้ดึงค่ารวม)
    def find_val(pattern):
        match = re.search(pattern, full_text, re.I)
        return float(match.group(1).replace(",", "")) if match else 0.0

    result["M"] = find_val(r'เงินค่าไฟฟ้าฐาน.*?([\d,]+\.\d+)')
    result["O"] = find_val(r'ค่า\s*Ft.*?=\s*[\d\.]+\s*บาท/หน่วย\s+([\d,]+\.\d+)')
    result["L"] = find_val(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)')
    result["Q"] = find_val(r'รวมเงินค่าไฟฟ้าเดือนปัจจุบัน\s*\(Total\)\s*([\d,]+\.\d+)')

    # 2. ดึงจากตาราง (สำหรับค่า Peak, PP, OP, H)
    for table in all_tables:
        for row in table:
            row_str = " ".join([str(cell) for cell in row if cell])
            
            # ตรรกะการเลือกค่า Peak/PP/OP
            if "พลังงาน" in row_str or "Peak" in row_str or "OP" in row_str:
                nums = [float(re.sub(r'[^\d.]', '', str(n))) for n in row if n and re.search(r'\d', str(n))]
                if len(nums) >= 1:
                    # ตัวอย่างการใส่ค่า (ปรับตามลำดับคอลัมน์ในตาราง PEA ของคุณ)
                    if "Peak" in row_str: result["I"] = nums[-1] # พลังงาน P
                    if "OP" in row_str: result["K"] = nums[-1]   # พลังงาน OP
                    if "H" in row_str: result["H"] = nums[-1]    # ค่า H

    return result

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

            # 1. ฟังก์ชัน write_number (กำหนดไว้ในนี้เพื่อให้เรียกใช้ได้ง่าย)
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

            # 2. ส่วนการวนลูปกรอกข้อมูล
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
