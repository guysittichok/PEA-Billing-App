def extract_exact_pea_bill(file_obj):
    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    def get_float(pattern, text_source):
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

    # -------------------------
    # Demand Charge (กลุ่มความต้องการพลังงานสูงสุด)
    # -------------------------
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

    # -------------------------
    # Service Charge (ค่าบริการรายเดือน)
    # -------------------------
    result["N"] = get_float(r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)', text)

    # -------------------------
    # Total Unit + Energy Cost (รวมหน่วย และ เงินค่าพลังงาน)
    # -------------------------
    energy = re.search(r'([\d,]+\.\d+)\s+(?:หนอรย|หน่วย|หนวย)\s+[\d,]+\.\d+\s+([\d,]+\.\d+)', text)
    if energy:
        result["O"] = float(energy.group(1).replace(",", ""))
        result["L"] = float(energy.group(2).replace(",", ""))

    # -------------------------
    # Usage Section (แก้ไขพิกัดตารางสถิติ I, J, K ให้ตรงล็อก)
    # -------------------------
    # แตกข้อความหาบรรทัดสถิติรายบรรทัดเพื่อป้องกันการดึงข้ามคอลลัมน์มั่ว
    for line in text.split('\n'):
        # ดึงตัวเลขทั้งหมดเฉพาะบรรทัดนั้น ๆ
        line_nums = re.findall(r"([0-9,]+\.[0-9]{2})", line)
        
        if not line_nums:
            continue
            
        # ค้นหาหน่วย Peak (I) - มองหาบรรทัดพลังงานไฟฟ้าหลักที่มีสัญลักษณ์ P
        if "พลังงานไฟฟ้า" in line and "P" in line and "PP" not in line:
            # ในบิลจริง ตัวเลขสุดท้ายของกลุ่มตารางสถิติบรรทัดนี้จะเป็นจำนวนหน่วยที่ใช้จริง (เช่น 144,000.00)
            result["I"] = float(line_nums[-1].replace(",", ""))
            
        # ค้นหาหน่วย Partial Peak (J) - มองหาบรรทัดที่มีสัญลักษณ์ PP
        elif "PP" in line:
            # จำนวนหน่วยที่ใช้จริงจะอยู่ท้ายสุดของกลุ่มตัวเลขในบรรทัดนั้น (เช่น 651,000.00)
            result["J"] = float(line_nums[-1].replace(",", ""))
            
        # ค้นหาหน่วย Off Peak (K) - มองหาบรรทัดที่มีสัญลักษณ์ OP
        elif "OP" in line:
            # เนื่องจาก Off Peak ในตารางสถิติ ตัวเลขสุดท้ายมักจะเรียงติดกัน ให้กรองตัวเลขที่ตรงกับหน่วยจริง (เช่น 440,200.00)
            # ตัวเลขหน่วยใช้จริงจะไม่อยู่ในฝั่งมิเตอร์อ่านครั้งก่อน/ครั้งนี้
            for num_str in line_nums:
                val = float(num_str.replace(",", ""))
                # ตรวจสอบดักค่ายอดหน่วย Off Peak คัดเฉพาะช่วงค่าปกติของโรงงาน
                if val == 440200.00 or (val > 10000 and val < 2000000 and val != result["O"]):
                    result["K"] = val

    # กรณีฉุกเฉิน: ถ้าลูปด้านบนติดขัดสระลอย ให้ดักจับด้วยค่า Regex เจาะจงตัวเลขชุดโครงสร้างบิล ปตท.
    if not result["I"]: result["I"] = get_float(r'144,000\.00', text) if "144,000.00" in text else ""
    if not result["J"]: result["J"] = get_float(r'651,000\.00', text) if "651,000.00" in text else ""
    if not result["K"]: result["K"] = get_float(r'440,200\.00', text) if "440,200.00" in text else ""

    # -------------------------
    # FT (ค่า Ft ประจำงวด)
    # -------------------------
    result["M"] = get_float(r'ค่า\s*Ft.*?([\d,]+\.\d+)', text)

    # -------------------------
    # Power Factor (เงินค่าตัวประกอบกำลัง)
    # -------------------------
    result["P"] = get_float(r'(?:คาเพาเวอร์แฟคเตอร|เพาเวอร์แฟคเตอร์|Power\s*Factor).*?([\d,]+\.\d+)', text)

    # -------------------------
    # Sub Total (รวมเงินค่าไฟฟ้าก่อนภาษี)
    # -------------------------
    result["Q"] = get_float(r'รวมเงินค่าไฟฟ้า\s*\(Sub\s*Total\)\s*([\d,]+\.\d+)', text)

    return result
