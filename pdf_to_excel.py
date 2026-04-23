"""
PDF Trade Report → Excel Portfolio Report converter
PASHA Kapital → PASHA Private Banking format

Usage:
    python pdf_to_excel.py <input.pdf> <template.xlsx> [output.xlsx]
"""

import sys
import re
import pdfplumber
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import os


# ── PDF extraction ──────────────────────────────────────────────────────────

SECURITY_ALIASES = {
    "SPDR S&P 500 ETF TRUST": "SPDR S&P 500 ETF TRUST (SPY)",
    "SPY": "SPDR S&P 500 ETF TRUST (SPY)",
    "SPDR GOLD": "SPDR GOLD SHARES (GLD)",
    "GLD": "SPDR GOLD SHARES (GLD)",
    "VANGUARD MID-CAP": "VANGUARD MID-CAP ETF (VO)",
    "MID-CAP ETF": "VANGUARD MID-CAP ETF (VO)",
    "ISHARES 7-10": "iShares 7-10 YEAR TREASURY BOND (IEF)",
    "7-10 YEAR TREASURY": "iShares 7-10 YEAR TREASURY BOND (IEF)",
    "ISHARES 1-5": "iSHARES 1-5Y INVESTMENT GRADE CORP BOND",
    "1-5Y": "iSHARES 1-5Y INVESTMENT GRADE CORP BOND",
    "INV GRADE CORP": "iSHARES 1-5Y INVESTMENT GRADE CORP BOND",
}

def clean_number(text):
    """Parse a number string like '(1,234.56)' or '31,083.76' → float."""
    if text is None:
        return None
    text = str(text).strip().replace(" ", "").replace(",", "")
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        val = float(text)
        return -val if negative else val
    except ValueError:
        return None


def normalize_security_name(raw):
    """Map raw PDF security name to the canonical Excel name."""
    raw_upper = raw.upper()
    for key, canonical in SECURITY_ALIASES.items():
        if key.upper() in raw_upper:
            return canonical
    return raw.strip()


def extract_pdf_data(pdf_path):
    """
    Extract all relevant numbers from the PASHA Kapital trade report PDF.
    Returns a dict with:
        holdings     : list of dicts per security
        commissions  : dict with broker/management totals
        report_date  : string
        client_name  : string
    """
    data = {
        "holdings": [],
        "commissions": {"broker_ccy": None, "broker_azn": None,
                        "management_ccy": None, "management_azn": None,
                        "total_ccy": None, "total_azn": None},
        "report_date": "",
        "client_name": "",
        "total_opening_ccy": None,
        "total_opening_azn": None,
        "total_closing_ccy": None,
        "total_closing_azn": None,
    }

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            tables = page.extract_tables()

            # ── Client / date from header text ──
            for line in text.splitlines():
                if "Müştərinin adı" in line or "Customer name" in line:
                    parts = line.split("/")
                    if len(parts) > 1:
                        data["client_name"] = parts[-1].strip()
                if re.search(r"\d{1,2}\s+\w+\s+20\d{2}", line):
                    m = re.search(r"(\d{1,2}[./]\d{1,2}[./]20\d{2})", line)
                    if m and not data["report_date"]:
                        data["report_date"] = m.group(1)

            # ── Parse tables ──
            for table in tables:
                if not table:
                    continue
                for row in table:
                    clean_row = [str(c).strip() if c else "" for c in row]
                    row_text = " ".join(clean_row).upper()

                    # Detect holdings rows — rows containing known security keywords
                    sec_name = None
                    for key in SECURITY_ALIASES:
                        if key.upper() in row_text:
                            sec_name = normalize_security_name(key)
                            break

                    if sec_name:
                        # Try to extract numeric columns from the row
                        numbers = []
                        for cell in clean_row:
                            n = clean_number(cell)
                            if n is not None:
                                numbers.append(n)

                        holding = {"name": sec_name, "numbers": numbers}

                        # Try to assign columns by position heuristics
                        # PDF main table order (approximate):
                        # lot/qty | price_type_cols... | opening_ccy | opening_azn | closing_ccy | closing_azn
                        if len(numbers) >= 4:
                            holding["quantity"] = int(numbers[0]) if numbers[0] == int(numbers[0]) else numbers[0]
                            # last 4 numbers: open_ccy, open_azn, close_ccy, close_azn
                            holding["opening_ccy"] = numbers[-4]
                            holding["opening_azn"] = numbers[-3]
                            holding["closing_ccy"] = numbers[-2]
                            holding["closing_azn"] = numbers[-1]
                        elif len(numbers) >= 2:
                            holding["opening_ccy"] = numbers[-2]
                            holding["closing_ccy"] = numbers[-1]

                        data["holdings"].append(holding)

                    # Detect commission summary rows
                    if "TOTAL" in row_text or "CƏMİ" in row_text or "CƏMI" in row_text:
                        nums = [clean_number(c) for c in clean_row if clean_number(c) is not None]
                        if len(nums) >= 2:
                            data["commissions"]["total_ccy"] = nums[-2]
                            data["commissions"]["total_azn"] = nums[-1]

    return data


# ── Excel update ─────────────────────────────────────────────────────────────

def find_cell(ws, search_text, search_cols=None, max_row=100):
    """Find first cell containing search_text (case-insensitive)."""
    for row in ws.iter_rows(min_row=1, max_row=max_row):
        for cell in row:
            if search_cols and cell.column not in search_cols:
                continue
            if cell.value and search_text.lower() in str(cell.value).lower():
                return cell
    return None


def update_excel(template_path, data, output_path):
    wb = openpyxl.load_workbook(template_path)

    # Work on the first sheet named FINAL (or first sheet)
    sheet_name = None
    for name in wb.sheetnames:
        if "FINAL" in name.upper():
            sheet_name = name
            break
    ws = wb[sheet_name] if sheet_name else wb.active

    print(f"  Updating sheet: {ws.title}")

    # ── Update Holdings table ──────────────────────────────────────────────
    # Find the "Name" header in the holdings table
    name_header = find_cell(ws, "Name")
    if name_header:
        name_col = name_header.column
        header_row = name_header.row

        # Read column headers in that row to find Quantity, Executed Price, etc.
        col_map = {}
        for cell in ws[header_row]:
            if cell.value:
                val = str(cell.value).lower()
                if "name" in val:
                    col_map["name"] = cell.column
                elif "quantity" in val or "qty" in val:
                    col_map["quantity"] = cell.column
                elif "executed" in val or "exec" in val:
                    col_map["executed_price"] = cell.column
                elif "trading" in val:
                    col_map["trading_value"] = cell.column
                elif "current price" in val:
                    col_map["current_price"] = cell.column
                elif "current value" in val:
                    col_map["current_value"] = cell.column

        print(f"  Found holdings header at row {header_row}, columns: {col_map}")

        # Match PDF holdings to Excel rows by security name
        for pdf_holding in data["holdings"]:
            # Scan rows below header to find matching security
            for r in range(header_row + 1, header_row + 20):
                cell_val = ws.cell(row=r, column=name_col).value
                if not cell_val:
                    continue
                cell_name = str(cell_val).strip().upper()
                pdf_name = pdf_holding["name"].upper()

                # Match if any significant word overlaps
                match = any(
                    word in cell_name
                    for word in pdf_name.split()
                    if len(word) > 3
                )

                if match:
                    print(f"    Matched: '{cell_val}' ← {pdf_holding['name']}")

                    if "quantity" in col_map and "quantity" in pdf_holding:
                        ws.cell(row=r, column=col_map["quantity"]).value = pdf_holding["quantity"]

                    if "executed_price" in col_map and "opening_ccy" in pdf_holding:
                        # Executed price = opening amount / quantity
                        qty = pdf_holding.get("quantity", 1) or 1
                        ws.cell(row=r, column=col_map["executed_price"]).value = round(
                            pdf_holding["opening_ccy"] / qty, 2
                        )

                    if "trading_value" in col_map and "opening_ccy" in pdf_holding:
                        ws.cell(row=r, column=col_map["trading_value"]).value = pdf_holding["opening_ccy"]

                    if "current_price" in col_map and "closing_ccy" in pdf_holding:
                        qty = pdf_holding.get("quantity", 1) or 1
                        ws.cell(row=r, column=col_map["current_price"]).value = round(
                            pdf_holding["closing_ccy"] / qty, 2
                        )

                    if "current_value" in col_map and "closing_ccy" in pdf_holding:
                        ws.cell(row=r, column=col_map["current_value"]).value = pdf_holding["closing_ccy"]

                    break
    else:
        print("  WARNING: Could not find 'Name' header in Excel — check the template sheet.")

    # ── Update commission / fee rows ──────────────────────────────────────
    if data["commissions"]["total_ccy"] is not None:
        mgmt_cell = find_cell(ws, "Management Fee")
        if mgmt_cell:
            # Fee is typically 2 columns to the right of the label
            fee_col = mgmt_cell.column + 2
            ws.cell(row=mgmt_cell.row, column=fee_col).value = -abs(data["commissions"]["total_ccy"])
            print(f"  Updated Management Fee: {data['commissions']['total_ccy']}")

    wb.save(output_path)
    print(f"\n  Saved: {output_path}")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 3:
        print("Usage: python pdf_to_excel.py <input.pdf> <template.xlsx> [output.xlsx]")
        print("\nExample:")
        print("  python pdf_to_excel.py trade_report.pdf portfolio_template.xlsx output.xlsx")
        sys.exit(1)

    pdf_path = sys.argv[1]
    template_path = sys.argv[2]

    # Default: save output next to the PDF file
    if len(sys.argv) > 3:
        output_path = sys.argv[3]
    else:
        pdf_dir = os.path.dirname(os.path.abspath(pdf_path))
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = os.path.join(pdf_dir, f"{pdf_name}_portfolio.xlsx")

    if not os.path.exists(pdf_path):
        print(f"ERROR: PDF not found: {pdf_path}")
        sys.exit(1)
    if not os.path.exists(template_path):
        print(f"ERROR: Template not found: {template_path}")
        sys.exit(1)

    print(f"Extracting data from: {pdf_path}")
    data = extract_pdf_data(pdf_path)

    print(f"\nExtracted {len(data['holdings'])} holdings:")
    for h in data["holdings"]:
        print(f"  {h['name']}: {h}")

    print(f"\nUpdating Excel template: {template_path}")
    update_excel(template_path, data, output_path)

    print("\nDone!")


if __name__ == "__main__":
    main()
