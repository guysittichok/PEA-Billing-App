import pdfplumber
import re

def extract_pea_bill(pdf_file):

    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(
            page.extract_text()
            for page in pdf.pages
            if page.extract_text()
        )

    def get(pattern):
        match = re.search(pattern, text, re.MULTILINE)
        if match:
            return float(match.group(1).replace(",", ""))
        return None

    data = {}

    # ------------------------
    # Customer Info
    # ------------------------

    ca_match = re.search(r'(\d{12})\s+838', text)
    if ca_match:
        data["CA_No"] = ca_match.group(1)

    # ------------------------
    # Demand
    # ------------------------

    data["Peak_Demand_kW"] = get(
        r'Peak\s+([\d,]+\.\d+)\s+กว'
    )

    data["PartialPeak_Demand_kW"] = get(
        r'Partial Peak\s+([\d,]+\.\d+)\s+กว'
    )

    data["OffPeak_Demand_kW"] = get(
        r'Off Peak\s+([\d,]+\.\d+)\s+กว'
    )

    # ------------------------
    # Demand Cost
    # ------------------------

    peak_cost = re.search(
        r'Peak\s+[\d,]+\.\d+\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text
    )

    if peak_cost:
        data["Peak_Demand_Cost"] = float(
            peak_cost.group(1).replace(",", "")
        )

    pp_cost = re.search(
        r'Partial Peak\s+[\d,]+\.\d+\s+กว\.\s+[\d,]+\.\d+\s+([\d,]+\.\d+)',
        text
    )

    if pp_cost:
        data["PartialPeak_Demand_Cost"] = float(
            pp_cost.group(1).replace(",", "")
        )

    # ------------------------
    # Energy Unit
    # ------------------------

    energy_match = re.search(
        r'พลังงานไฟฟ้า.*?P.*?([\d,]+\.\d+).*?PP.*?([\d,]+\.\d+).*?OP.*?([\d,]+\.\d+)',
        text,
        re.S
    )

    if energy_match:

        data["Peak_Energy_Unit"] = float(
            energy_match.group(1).replace(",", "")
        )

        data["PartialPeak_Energy_Unit"] = float(
            energy_match.group(2).replace(",", "")
        )

        data["OffPeak_Energy_Unit"] = float(
            energy_match.group(3).replace(",", "")
        )

    # ------------------------
    # Summary
    # ------------------------

    data["Based_Amount"] = get(
        r'เงินค่าไฟฟ้าฐาน \(Based Amount\)\s+([\d,]+\.\d+)'
    )

    data["Ft"] = get(
        r'ค่า Ft .*?([\d,]+\.\d+)'
    )

    data["Power_Factor"] = get(
        r'ค่าเพาเวอร์แฟคเตอร์\s+([\d,]+\.\d+)'
    )

    data["Sub_Total"] = get(
        r'รวมเงินค่าไฟฟ้า \(Sub Total\)\s+([\d,]+\.\d+)'
    )

    data["VAT"] = get(
        r'ภาษีมูลค่าเพิ่ม.*?([\d,]+\.\d+)'
    )

    data["Total"] = get(
        r'รวมเงินค่าไฟฟ้าเดือนปัจจุบัน \(Total\)\s+([\d,]+\.\d+)'
    )

    return data
