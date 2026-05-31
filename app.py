def extract_exact_pea_bill(file_obj):

    with pdfplumber.open(file_obj) as pdf:
        text = "\n".join(
            page.extract_text()
            for page in pdf.pages
            if page.extract_text()
        )

    def get_float(pattern):
        m = re.search(pattern, text, re.S | re.I)
        if m:
            return float(m.group(1).replace(",", ""))
        return ""

    result = {
        "ชื่อไฟล์": file_obj.name,

        "C": "",
        "D": "",
        "E": "",
        "F": "",
        "G": "",
        "H": "",
        "I": "",
        "J": "",
        "K": "",
        "L": "",
        "M": "",
        "N": "",
        "O": "",
        "P": "",
        "Q": ""
    }

    # -------------------------
    # Demand Charge
    # -------------------------

    peak = re.search(
        r'Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text,
        re.I
    )

    if peak:
        result["C"] = float(peak.group(1).replace(",", ""))
        result["F"] = float(peak.group(2).replace(",", ""))

    pp = re.search(
        r'Partial Peak\s+([\d,]+\.\d+)\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text,
        re.I
    )

    if pp:
        result["D"] = float(pp.group(1).replace(",", ""))
        result["G"] = float(pp.group(2).replace(",", ""))

    op = re.search(
        r'Off Peak\s+([\d,]+\.\d+)\s+กว',
        text,
        re.I
    )

    if op:
        result["E"] = float(op.group(1).replace(",", ""))

    # -------------------------
    # Service Charge
    # -------------------------

    result["N"] = get_float(
        r'ค่าบริการรายเดือน.*?([\d,]+\.\d+)'
    )

    # -------------------------
    # Total Unit + Energy Cost
    # -------------------------

    energy = re.search(
        r'([\d,]+\.\d+)\s+หน่วย\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text
    )

    if energy:
        result["O"] = float(energy.group(1).replace(",", ""))
        result["L"] = float(energy.group(2).replace(",", ""))

    # -------------------------
    # Usage Section
    # -------------------------

    usage_peak = re.search(
        r'พลังงานไฟฟ้า.*?P\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text,
        re.S
    )

    if usage_peak:
        result["I"] = float(
            usage_peak.group(1).replace(",", "")
        )

    usage_pp = re.search(
        r'PP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text
    )

    if usage_pp:
        result["J"] = float(
            usage_pp.group(1).replace(",", "")
        )

    usage_op = re.search(
        r'OP\s+[\d,]+\.\d+\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text
    )

    if usage_op:
        result["K"] = float(
            usage_op.group(1).replace(",", "")
        )

    # -------------------------
    # FT
    # -------------------------

    result["M"] = get_float(
        r'ค่า Ft.*?([\d,]+\.\d+)'
    )

    # -------------------------
    # Power Factor
    # -------------------------

    result["P"] = get_float(
        r'ค่าเพาเวอร์แฟคเตอร์\s+([\d,]+\.\d+)'
    )

    # -------------------------
    # Sub Total
    # -------------------------

    result["Q"] = get_float(
        r'รวมเงินค่าไฟฟ้า \(Sub Total\)\s+([\d,]+\.\d+)'
    )

    return result
