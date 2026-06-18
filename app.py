import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
import openpyxl
# 🛠️ นำเข้าโมดูลจัดการ สี, ฟอนต์ และการจัดวาง ของ openpyxl
from openpyxl.styles import Font, Alignment

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

    # ตรวจสอบรูปแบบบิล TOU
    is_tou = any(re.search(r''+k+r'.*?(?:กว|หน่วย|หนอรย|หนวย)', text, re.I) for k in ["Peak", "Off Peak", "Partial Peak"]) or ("OP" in text and " P " in text)
    has_h_mode = " H " in text or "\nH " in text or " H\n" in text or " Holiday " in text

    # ========================================================
    # บล็อกลอจิกเดิมทุกช่อง (ห้ามแตะต้องเด็ดขาด)
    # ========================================================
    if is_tou and has_h_mode:
        peak_money_match = re.search(r'Peak\s+[\d,]+\.\d+\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
        if peak_money_match: result["F"] = float(peak_money_match.group(1).replace(",", ""))

        if result["F"] == 0.0 or result["F"] == 0:
            gwa_pattern = re.search(r'([\d,]+\.\d+)\s+กว\.?\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
            if gwa_pattern:
                result["F"] = float(gwa_pattern.group(2).replace(",", ""))
            
        off_peak_money_match = re.search(r'Off\s+Peak\s+[\d,]+\.\d+\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
        if off_peak_money_match: result["G"] = float(off_peak_money_match.group(1).replace(",", ""))
            
        h_money_match = re.search(r'(?:H|Holiday)\s+[\d,]+\.\d+\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
        if h_money_match: result["H"] = float(h_money_match.group(1).replace(",", ""))

        parts = text.split("พลังงานไฟฟ้า")
        demand_part = parts[0]
        energy_part = parts[1] if len(parts) > 1 else text 

        demand_lines = demand_part.split('\n')
        for line in demand_lines:
            nums = re.findall(r"([\d,]+\.\d+)", line)
            if not nums: continue
            if "P" in line and "กว" in line and "Off" not in line:
                result["C"] = float(nums[2].replace(",", "")) if len(nums) >= 3 else float(nums[0].replace(",", ""))
            elif "OP" in line:
                result["D"] = float(nums[2].replace(",", "")) if len(nums) >= 3 else float(nums[0].replace(",", ""))
            elif line.strip().startswith("H ") or " H " in line or "Holiday" in line:
                if "กว" in line or len(nums) >= 3:
                    result["E"] = float(nums[2].replace(",", "")) if len(nums) >= 3 else float(nums[0].replace(",", ""))

        p_unit = re.search(r'(?:^|\s+)P\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', energy_part, re.I)
        if p_unit: result["I"] = float(p_unit.group(1).replace(",", ""))
        else:
            p_unit_alt = re.search(r'(?:พลังงานไฟฟ้า)?\s+P\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text, re.I)
            if p_unit_alt: result["I"] = float(p_unit_alt.group(1).replace(",", ""))

        op_unit = re.search(r'OP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', energy_part, re.I)
        if op_unit: result["J"] = float(op_unit.group(1).replace(",", ""))

        h_unit = re.search(r'(?:H|Holiday)\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', energy_part, re.I)
        if h_unit: result["K"] = float(h_unit.group(1).replace(",", ""))
            
        if result["K"] == 0.0:
            h_fallback = re.search(r'OP.*?([\d,]+\.\d+)\s+(?:H|Holiday)\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', energy_part, re.DOTALL | re.I)
            if h_fallback: result["K"] = float(h_fallback.group(2).replace(",", ""))

    elif not is_tou:
        lines = text.split('\n')
        unit_match = re.search(r'([\d,]+\.\d+)\s+(?:หน่วย|หนอรย|หนวย)', text)
        if unit_match: result["I"] = float(unit_match.group(1).replace(",", ""))

        if "พลังไฟฟ้าสูงสุด" in text:
            for line in lines:
                if "พลังไฟฟ้าสูงสุด" in line:
                    nums = re.findall(r"([\d,]+\.\d+)", line)
                    if nums:
                        if len(nums) >= 3: result["C"] = float(nums[2].replace(",", ""))
                        else: result["C"] = float(nums[-1].replace(",", ""))

            demand_cost_match = re.search(r'พลังไฟฟ้าสูงสุด\s+.*?กว\..*?([\d,]+\.\d+)', text)
            if demand_cost_match: result["F"] = float(demand_cost_match.group(1).replace(",", ""))

            if result["F"] == 0.0 or result["F"] == 0:
                gwa_pattern = re.search(r'([\d,]+\.\d+)\s+กว\.\s+([\d,]+\.\d+)\s+([\d,]+\.\d+)', text)
                if gwa_pattern:
                    result["C"] = float(gwa_pattern.group(1).replace(",", ""))
                    result["F"] = float(gwa_pattern.group(3).replace(",", ""))
            if result["F"] == 0.0 or result["F"] == 0:
                last_ditch_match = re.findall(r'กว\..*?([\d,]+\.\d+)', text)
                if last_ditch_match:
                    result["F"] = float(last_ditch_match[-1].replace(",", ""))
            
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

    # ดึงค่าช่อง M (Off Peak หน่วยสะสมซ้ายสุด)
    for line in text.split('\n'):
        if "off" in line.lower() and "peak" in line.lower() and any(k in line for k in ["หน่วย", "หนอรย", "หนวย"]):
            clean_line = re.split(r'\d{2}/\d{2}/\d{2,4}', line)[0]
            nums_in_op_line = re.findall(r"([\d,]+\.\d+)", clean_line)
            if nums_in_op_line:
                result["M"] = float(nums_in_op_line[-1].replace(",", ""))
                break

    # ========================================================
    # ระบบ Block Search Pattern (ช่อง L)
    # ========================================================
    found_l = False

    if is_tou:
        tou_pattern = re.search(r'(?:Peak|[\d,]+\.\d+)\s+(?:หน่วย|หนอรย|หนวย)\s+(\d+\.\d{4})\s+([\d,]+\.\d+)', text, re.I)
        if tou_pattern:
            potential_money = float(tou_pattern.group(2).replace(",", ""))
            if potential_money != result["M"] and potential_money not in [65760.0, 379280.0]:
                result["L"] = potential_money
                found_l = True
        
        if not found_l:
            for line in text.split('\n'):
                if re.search(r'\d{2}/\d{2}/\d{2,4}', line) or "ประวัติ" in line or "history" in line.lower() or "กว" in line:
                    continue
                if any(k in line for k in ["หน่วย", "หนอรย", "หนวย"]):
                    nums = re.findall(r"([\d,]+\.\d+)", line)
                    if len(nums) >= 3:
                        if any(len(n.split('.')[-1]) == 4 for n in nums):
                            potential_money = float(nums[-1].replace(",", ""))
                            if potential_money != result["M"] and potential_money not in [65760.0, 379280.0]:
                                result["L"] = potential_money
                                found_l = True
                                break

    if not is_tou or not found_l:
        for line in text.split('\n'):
            if re.search(r'\d{2}/\d{2}/\d{2,4}', line) or any(k in line for k in ["ประวัติ", "history", "กว.", "กว", "ค่าบริการ", "รวมเงิน", "total"]): 
                continue
            if any(k in line for k in ["หน่วย", "หนอรย", "หนวย"]):
                nums_in_line = re.findall(r"([\d,]+\.\d+)", line)
                if nums_in_line:
                    result["L"] = float(nums_in_line[-1].replace(",", ""))
                    break

    # ========================================================
    # ดึงค่า Ft, PF และ Sub Total
    # ========================================================
    ft = re.search(r'ค่า\s*Ft.*?=\s*[\d\.]+\s*บาท/หน่วย\s+([\d,]+\.\d+)', text, re.I)
    if not ft: ft = re.search(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text, re.I)
    if ft: result["O"] = float(ft.group(1).replace(",", ""))
    
    result["P"] = 0.0  
    for line in text.split('\n'):
        if any(k in line for k in ["ค่าเพาเวอร์แฟคเตอร", "เพาเวอร์แฟคเตอร์", "Power Factor", "คาเพาเวอร์แฟคเตอร"]):
            nums_in_pf_line = re.findall(r"([\d,]+\.\d+)", line)
            if nums_in_pf_line:
                result["P"] = float(nums_in_pf_line[-1].replace(",", ""))
                break 

    total_match = re.search(r'รวมเงินค่าไฟฟ้า\s*\(Sub Total\)\s*([\d,]+\.\d+)', text, re.I)
    if total_match: result["Q"] = float(total_match.group(1).replace(",", ""))

    return result

# ----------------------------------------------------
# ส่วนโครงสร้าง Excel & Streamlit UI 
# ----------------------------------------------------
uploaded_files = st.file_uploader("อัปโหลดไฟล์บิล PDF", type=["pdf"], accept_multiple_files=True)
template_file = st.file_uploader("อัปโหลดไฟล์ Excel ต้นแบบ", type=["xlsx"])

if uploaded_files:
    data = [extract_exact_pea_bill(f) for f in uploaded_files]
    all_cols = ["ชื่อไฟล์", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"]
    df = pd.DataFrame(data, columns=all_cols)
    
    df["O"] = "-"
    df["Q"] = "-"
    
    st.data_editor(df, use_container_width=True)

    if template_file and st.button("สร้างไฟล์ Excel พร้อมข้อมูล"):
        try:
            wb = openpyxl.load_workbook(template_file)
            
            # 🎯 แก้ไขลอจิกการดึงหน้าชีท เพื่อป้องกันการเกิดบั๊ก NoneType 
            if wb.active is not None:
                ws = wb.active
            else:
                ws = wb.worksheets[0]

            # 🎯 ปรับแต่งฟอนต์สีน้ำเงินแบบ "ตัวปกติ (bold=False)" และชิดขวามือทั้งหมด
            blue_normal_font = Font(name="Calibri", size=11, bold=False, color="0000FF")
            right_alignment = Alignment(horizontal="right", vertical="center")

            def write_number(ws, cell_pos, value):
                val_str = str(value).strip()
    
                # บังคับใช้สีฟอนต์ตัวปกติสีน้ำเงิน และ จัดตำแหน่งชิดขวา
                ws[cell_pos].font = blue_normal_font
                ws[cell_pos].alignment = right_alignment
    
                # กำหนดรูปแบบ Number Format แบบ Accounting (ถ้าเป็น 0 จะแสดงผลเป็นเครื่องหมายขีด - อัตโนมัติ)
                # และสามารถนำไป บวก ลบ คูณ หาร ใน Excel ได้ปกติ ไม่เกิด #VALUE!
                accounting_format = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'
    
                if val_str in ["None", "", "-"]:
                    ws[cell_pos] = 0
                    ws[cell_pos].number_format = accounting_format
                else:
                    try:
                        val = float(val_str.replace(',', ''))
                        ws[cell_pos] = val
                        ws[cell_pos].number_format = accounting_format
                    except:
                        # หากแปลงค่าไม่ได้จริงๆ ให้ใส่เลข 0 ไว้ก่อนเพื่อความปลอดภัยในสูตรคำนวณ
                        ws[cell_pos] = 0
                        ws[cell_pos].number_format = accounting_format

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
                    write_number(ws, f'P{row_idx}', row['P'])
            
            output = BytesIO()
            
            # 🎯 เพิ่มลอจิกเปิดโหมดบังคับคำนวณสูตรออโต้ เพื่อไม่ให้ค่าในสูตรแสดงเป็นช่องว่างตอนดาวน์โหลดออกไปครับ
            wb.properties.calcMode = 'auto'
            
            wb.save(output)
            output.seek(0)
            st.success("กรอกข้อมูลลงใน Template เรียบร้อยและรักษาสูตรเดิมแล้ว!")
            st.download_button("📥 ดาวน์โหลด Excel", output, "Updated_PEA_Bill.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# ฟังก์ชันช่วยจัดรูปแบบข้อความและฟอนต์ไทย
def add_run_thai(paragraph, text, font_name="TH Sarabun PSK", size_pt=16, bold=False, color_rgb=None):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    # บังคับอักษรภาษาไทย/Complex Script ใน XML ของ Word
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)
    return run

# ฟังก์ชันระบายสีพื้นหลังเซลล์ตาราง (Shading)
def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

# ฟังก์ชันลบเส้นขอบตารางเฉพาะจุด (สำหรับตาราง Header)
def set_table_borders_horizontal_only(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="none"/>'
        f'  <w:left w:val="none"/>'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CCCCCC"/>'
        f'  <w:right w:val="none"/>'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E0E0"/>'
        f'  <w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def create_memo_report(df_data, selected_month, selected_year):
    doc = Document()
    
    # ตั้งค่าหน้ากระดาษ ด้านละ 1 นิ้วตามมาตรฐาน
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # --- ส่วนหัวเอกสาร (Header) แยก ซ้าย - ขวา ---
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # ฝั่งซ้าย: ใส่โลโก้ PTT 
    cell_left = header_table.cell(0, 0)
    p_logo = cell_left.paragraphs[0]
    try:
        p_logo.add_run().add_picture("ptt_logo.png", width=Inches(1.6))
    except:
        add_run_thai(p_logo, "บริษัท ปตท. จำกัด (มหาชน)\nPTT Public Company Limited", size_pt=14, bold=True)
    
    # ฝั่งขวา: ใส่แถบกล่องดำ MEMORANDUM
    cell_right = header_table.cell(0, 1)
    p_memo = cell_right.paragraphs[0]
    p_memo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # ทำแถบดำจำลองความเป๊ะของ MEMORANDUM
    memo_table = cell_right.add_table(rows=1, cols=1)
    memo_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    memo_cell = memo_table.cell(0, 0)
    set_cell_background(memo_cell, "1A1A1A") # สีดำเข้ม
    memo_cell.width = Inches(2.2)
    p_box = memo_cell.paragraphs[0]
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_thai(p_box, "MEMORANDUM", size_pt=16, bold=True, color_rgb=RGBColor(255, 255, 255))

    doc.add_paragraph() # เว้นบรรทัดสั้นๆ

    # --- ส่วนตารางข้อมูลบันทึกข้อความ (ลบเส้นข้างออกตามต้นแบบ) ---
    info_table = doc.add_table(rows=5, cols=1)
    set_table_borders_horizontal_only(info_table)
    
    # บรรทัดที่ 1: ที่ / No และ วันที่ / Date
    p1 = info_table.cell(0, 0).paragraphs[0]
    add_run_thai(p1, "ที่ / No: ", bold=True)
    add_run_thai(p1, "-\t\t\t\t\t\t\t")
    add_run_thai(p1, "วันที่ / Date: ", bold=True)
    today_str = datetime.datetime.now().strftime(f"%d {selected_month} {selected_year}")
    add_run_thai(p1, today_str)

    # บรรทัดที่ 2: หน่วยงานผู้ส่ง
    p2 = info_table.cell(1, 0).paragraphs[0]
    add_run_thai(p2, "หน่วยงานผู้ส่ง / From: ", bold=True)
    add_run_thai(p2, "ส่วนบริหารกลยุทธ์และแผนการผลิต (กผ.)")

    # บรรทัดที่ 3: เรียน
    p3 = info_table.cell(2, 0).paragraphs[0]
    add_run_thai(p3, "เรียน / To: ", bold=True)
    add_run_thai(p3, "ผจ.สทต. / ผจ.บท ผ่าน ผจ.กผ.")

    # บรรทัดที่ 4: สำเนา
    p4 = info_table.cell(3, 0).paragraphs[0]
    add_run_thai(p4, "สำเนา / CC: ", bold=True)
    add_run_thai(p4, "ผจ.ปท.3, ผจ.ชก., ผจ.บฟ.")

    # บรรทัดที่ 5: เรื่อง
    p5 = info_table.cell(4, 0).paragraphs[0]
    add_run_thai(p5, "เรื่อง / Subject: ", bold=True)
    add_run_thai(p5, f"แจ้งสรุปค่าไฟฟ้าofสถานีชายฝั่งระยอง ประจำเดือน {selected_month} {selected_year}")

    doc.add_paragraph() # บรรทัดว่าง

    # --- ส่วนเนื้อความตอนต้น ---
    p_body1 = doc.add_paragraph()
    p_body1.paragraph_format.first_line_indent = Inches(0.5)
    p_body1.paragraph_format.line_spacing = 1.15
    add_run_thai(p_body1, f"ส่วนบริหารกลยุทธ์และแผนการผลิต ฝ่ายบริหารเทคนิคและแผนการผลิต ขอนำส่งสรุปค่าไฟฟ้าของสถานีชายฝั่งระยอง ประจำเดือน {selected_month} {selected_year} รายละเอียดการคำนวณตามเอกสารแนบ")

    # --- ส่วนตารางข้อมูลสรุปสีส้มพีช (สร้างตามภาพ Excel ต้นแบบ) ---
    # ใช้ค่า Default นิ่งๆ ตามที่คุณส่งมาเพื่อให้ใกล้เคียงที่สุดก่อนครับ
    table_data = [
        ["DPCU", "26,218.00", "121,418.83", "-", "-", "26,218.00", "121,418.83"],
        ["New DPCU", "60,451.30", "279,957.51", "-", "-", "60,451.30", "279,957.51"],
        ["OCS1", "132,224.00", "612,345.83", "-", "-", "132,224.00", "612,345.83"],
        ["OCS2", "111,774.00", "517,639.33", "-", "-", "111,774.00", "517,639.33"],
        ["OCS3", "421,808.00", "1,950,445.75", "-", "-", "421,808.00", "1,950,445.75"],
        ["Total", "752,475.30", "3,481,807.25", "-", "-", "752,475.30", "3,481,807.25"]
    ]

    calc_table = doc.add_table(rows=3, cols=7)
    calc_table.style = 'Table Grid'
    calc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # จัดการส่วนหัวตาราง (Header Merge)
    hdr_cells = calc_table.rows[0].cells
    hdr_cells_l2 = calc_table.rows[1].cells
    hdr_cells_l3 = calc_table.rows[2].cells
    
    # ระบายสีส้มพาสเทล/พีชที่หัวตาราง (#F8CECC หรือ #FCE4D6 ตามสไตล์ Excel)
    peach_color = "FCE4D6"
    
    # Merge หัวข้อ "พื้นที่ใช้ไฟฟ้า" (แถว 1-3 ยุบรวมกันคอลัมน์แรก)
    c0 = hdr_cells[0]
    c0.merge(hdr_cells_l2[0]).merge(hdr_cells_l3[0])
    set_cell_background(c0, peach_color)
    p_c0 = c0.paragraphs[0]
    p_c0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_thai(p_c0, "พื้นที่ใช้ไฟฟ้า", size_pt=13, bold=True)
    
    # Merge "ไฟฟ้าที่รับจาก PEA"
    c1 = hdr_cells[1]
    c1.merge(hdr_cells[2])
    set_cell_background(c1, peach_color)
    p_c1 = c1.paragraphs[0]
    p_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_thai(p_c1, "ไฟฟ้าที่รับจาก PEA", size_pt=13, bold=True)
    
    # Merge "ไฟฟ้าที่รับจาก GSP"
    c3 = hdr_cells[3]
    c3.merge(hdr_cells[4])
    set_cell_background(c3, peach_color)
    p_c3 = c3.paragraphs[0]
    p_c3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_thai(p_c3, "ไฟฟ้าที่รับจาก GSP", size_pt=13, bold=True)
    
    # Merge "รวมไฟฟ้าที่รับทั้งหมด"
    c5 = hdr_cells[5]
    c5.merge(hdr_cells[6])
    set_cell_background(c5, peach_color)
    p_c5 = c5.paragraphs[0]
    p_c5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run_thai(p_c5, "รวมไฟฟ้าที่รับทั้งหมด", size_pt=13, bold=True)
    
    # ใส่หัวข้อย่อยแถวที่ 2 และ 3
    sub_headers = ["ปริมาณไฟฟ้า", "ค่าใช้จ่าย", "ปริมาณไฟฟ้า", "ค่าใช้จ่าย", "ปริมาณไฟฟ้า", "ค่าใช้จ่าย"]
    sub_units = ["(kWh)", "(บาท)", "(kWh)", "(บาท)", "(kWh)", "(บาท)"]
    
    for idx in range(6):
        cell_h2 = hdr_cells_l2[idx+1]
        cell_h3 = hdr_cells_l3[idx+1]
        set_cell_background(cell_h2, peach_color)
        set_cell_background(cell_h3, peach_color)
        
        p_h2 = cell_h2.paragraphs[0]
        p_h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_thai(p_h2, sub_headers[idx], size_pt=11, bold=True)
        
        p_h3 = cell_h3.paragraphs[0]
        p_h3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run_thai(p_h3, sub_units[idx], size_pt=11, bold=True)

    # กรอกข้อมูลลงในตาราง
    for row_data in table_data:
        row_cells = calc_table.add_row().cells
        is_total = (row_data[0] == "Total")
        
        for idx, val in enumerate(row_data):
            cell = row_cells[idx]
            p = cell.paragraphs[0]
            
            # ถ้าเป็นแถว Total ให้ใส่พื้นหลังสีส้มพีชและตัวหนา
            if is_total:
                set_cell_background(cell, peach_color)
                
            # จัดตำแหน่งข้อความ: คอลัมน์แรกชิดกลาง คอลัมน์ที่เหลือชิดขวา
            if idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_run_thai(p, val, size_pt=12, bold=is_total)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                add_run_thai(p, val, size_pt=12, bold=is_total)

    doc.add_paragraph() # บรรทัดว่างหลังตาราง

    # --- ส่วนเนื้อความตอนท้าย (ดึงราคารวมจากตารางมาหยอดอัตโนมัติ) ---
    p_body2 = doc.add_paragraph()
    p_body2.paragraph_format.line_spacing = 1.15
    
    # ดึงค่าราคารวมจากตารางแถวสุดท้าย (คอลัมน์สุดท้ายของ Total)
    total_cost_str = table_data[-1][-1] 
    
    add_run_thai(p_body2, "\tอัตราค่าไฟฟ้าอ้างอิง ราคา PEA Rate ณ เดือน ")
    add_run_thai(p_body2, f"{selected_month} {selected_year} = ")
    add_run_thai(p_body2, "4.6311", bold=True)
    add_run_thai(p_body2, " บาท / kWh  รายละเอียดตามเอกสารแนบ\n")
    
    add_run_thai(p_body2, "\tรวมค่าไฟฟ้าที่เรียกเก็บจากระบบท่อส่งก๊าซฯ ทั้งสิ้น\t")
    add_run_thai(p_body2, f"{total_cost_str}", bold=True)
    add_run_thai(p_body2, "    \tบาท\n\n")
    add_run_thai(p_body2, "จึงเรียนมาเพื่อโปรดทราบ")

    doc.add_paragraph()
    
    # --- ส่วนลงชื่อท้ายประโยค ---
    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run_thai(p_sign, "(นางสุรีย์พันธ์ คุณากูลสวัสดิ์)\nผู้จัดการทั่วไป      ")

    # --- ส่วนเอกสารแนบ ---
    doc.add_paragraph()
    p_attach = doc.add_paragraph()
    add_run_thai(p_attach, "เอกสารแนบที่ 1 ", bold=True)
    add_run_thai(p_attach, "รายละเอียดในการคำนวณค่าไฟฟ้า\n")
    add_run_thai(p_attach, "เอกสารแนบที่ 2 ", bold=True)
    add_run_thai(p_attach, "หนังสือแจ้งค่าไฟฟ้าของการไฟฟ้าส่วนภูมิภาค")

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# เพิ่มส่วนควบคุมบน Streamlit UI (เมื่อมีการอัปโหลดไฟล์เสร็จแล้ว)
if uploaded_files:
    st.markdown("---")
    st.subheader("📊 ระบบออกรายงานสรุปบันทึกข้อความ (Memo Report)")
    
    col1, col2 = st.columns(2)
    with col1:
        months_list = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        selected_month = st.selectbox("เลือกประจำเดือน", months_list, index=3)
    with col2:
        years_list = [str(y) for y in range(2565, 2575)]
        selected_year = st.selectbox("เลือกปี พ.ศ.", years_list, index=4)
        
    if st.button("📝 สร้างรายงาน Word พร้อมตารางสรุป"):
        try:
            word_output = create_memo_report(df, selected_month, selected_year)
            st.success(f"สร้างบันทึกข้อความและตารางสรุปประจำเดือน {selected_month} {selected_year} สำเร็จ!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Memo (.docx)",
                data=word_output,
                file_name=f"Memo_สรุปค่าไฟฟ้า_{selected_month}_{selected_year}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการสร้างไฟล์: {e}")
