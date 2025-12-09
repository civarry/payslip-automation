"""PDF generation for payslips using ReportLab"""

import os
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from textwrap import wrap


def get_safe(row, col, default=0):
    """
    Safely get value from DataFrame row

    Args:
        row: pandas Series with employee data
        col: Column name to retrieve
        default: Default value if column doesn't exist

    Returns:
        Value from row or default
    """
    try:
        return row[col]
    except KeyError:
        return default


def create_payslip_pdf(row, output_dir, logo_path=None, company_config=None):
    """
    Generate a payslip PDF for a single employee

    Args:
        row: pandas Series with employee data
        output_dir: Directory to save the PDF
        logo_path: Path to company logo image (optional)
        company_config: Dict with company details (optional):
            - company_name: str
            - footer_text: str
            - document_id: str
            - effectivity_date: str

    Returns:
        str: Path to generated PDF file
    """
    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Extract employee data
    emp_no = str(row["EmployeeNumber"])
    name = str(row["Name"])
    period = str(row["PayrollPeriod"])
    position = str(row.get("Position", ""))

    # Create safe filename
    safe_period = period.replace(" ", "_").replace("/", "-").replace(",", "")
    filename = f"payslip_{emp_no}_{safe_period}.pdf"
    file_path = os.path.join(output_dir, filename)

    # Get company configuration or use defaults
    if company_config is None:
        company_config = {}

    company_name = company_config.get("company_name", "MASSPOWER PHILIPPINES ELECTRONIC INC.")
    footer_text = company_config.get("footer_text",
        "Full details of your pay for the period covered are given above. "
        "Please check carefully. Any questions or discrepancy concerning the "
        "accuracy of this statement should be taken up with this Office immediately.")
    document_id = company_config.get("document_id", "D-MPFA-20004.02")
    effectivity_date = company_config.get("effectivity_date", "January 20, 2024")

    # ---------- LANDSCAPE PAGE ----------
    page = landscape(A4)
    c = canvas.Canvas(file_path, pagesize=page)
    width, height = page

    left = 25
    right = width - 25
    top = height - 40
    line_h = 18

    y = top

    # ---------- HEADER ----------
    if logo_path and os.path.exists(logo_path):
        logo_w = 100
        logo_h = 50
        c.drawImage(
            logo_path,
            (width - logo_w) / 2,
            y - logo_h,
            logo_w,
            logo_h,
            preserveAspectRatio=True,
            mask="auto",
        )
        y -= logo_h + 20

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, company_name)
    y -= 30

    # ---------- TOP INFO TABLE ----------
    table_top = y
    table_bottom = table_top - (line_h * 3)
    mid_x = (left + right) / 2

    c.rect(left, table_bottom, right - left, table_top - table_bottom)
    c.line(mid_x, table_bottom, mid_x, table_top)
    c.line(left, table_top - line_h, right, table_top - line_h)
    c.line(left, table_top - 2 * line_h, right, table_top - 2 * line_h)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(left + 8, table_top - line_h + 4, "Employee Number")
    c.drawString(mid_x + 8, table_top - line_h + 4, "Payroll Period")

    c.drawString(left + 8, table_top - 2 * line_h + 4, "Basic")
    c.drawString(mid_x + 8, table_top - 2 * line_h + 4, "Name")

    c.drawString(left + 8, table_top - 3 * line_h + 4, "Monthly Allowance")
    c.drawString(mid_x + 8, table_top - 3 * line_h + 4, "Department/Position")

    c.setFont("Helvetica", 10)
    c.drawString(left + 140, table_top - line_h + 4, emp_no)
    c.drawString(mid_x + 140, table_top - line_h + 4, period)

    c.drawString(left + 140, table_top - 2 * line_h + 4, f"{get_safe(row,'BasicSalary',0):,.2f}")
    c.drawString(mid_x + 140, table_top - 2 * line_h + 4, name)

    c.drawString(left + 140, table_top - 3 * line_h + 4, f"{get_safe(row,'Allowance',0):,.2f}")
    c.drawString(mid_x + 140, table_top - 3 * line_h + 4, position)

    y = table_bottom - 30

    # ---------- EARNINGS & DEDUCTIONS TABLES ----------
    mid_table = (left + right) / 2
    earnings_left = left
    earnings_right = mid_table - 20
    ded_left = mid_table + 20
    ded_right = right

    rows_earn = [
        ("Regular Hours", "RegularHours", "RegularAmount"),
        ("Regular OT", "RegularOTHours", "RegularOTAmount"),
        ("Legal Holiday", "LegalHolidayHours", "LegalHolidayAmount"),
        ("Legal Holiday OT", None, None),
        ("Special Holiday", "SpecialHolidayHours", "SpecialHolidayAmount"),
        ("Special Holiday OT", None, None),
        ("Total Night Diff.", "NightDiffHours", "NightDiffAmount"),
        ("Offset", "OffsetHours", "OffsetAmount"),
        ("Paid Leave", None, "PaidLeaveAmount"),
        ("Adjustment", None, "AdjustmentEarnings"),
        ("Allowance", None, "Allowance"),
        ("13th Month Pay", None, "ThirteenthMonthPay"),
        ("Others", None, "OthersEarnings"),
        ("Gross Income", None, "GrossIncome"),
    ]

    rows_ded = [
        ("Pag-ibig Contribution", "PagibigContribution"),
        ("Philhealth Contribution", "PhilhealthContribution"),
        ("SSS Contribution", "SSSContribution"),
        ("Pag-ibig Loan", "PagibigLoan"),
        ("SSS Loan", "SSSLoan"),
        ("Withholding Tax", "WithholdingTax"),
        ("Adjustment", "AdjustmentDeductions"),
        ("Others", "OthersDeductions"),
        ("Total Deductions", "TotalDeductions"),
        ("NET PAY", "NetPay"),  # inside the table
    ]

    # ----- EARNINGS BOX -----
    earn_height = line_h * (len(rows_earn) + 1)
    earn_bottom = y - earn_height
    c.rect(earnings_left, earn_bottom, earnings_right - earnings_left, earn_height)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(earnings_left + 5, y - line_h + 4, "EARNINGS")
    c.drawString(earnings_left + 160, y - line_h + 4, "HOURS")
    c.drawString(earnings_left + 250, y - line_h + 4, "AMOUNT")
    c.line(earnings_left, y - line_h, earnings_right, y - line_h)

    c.setFont("Helvetica", 10)
    y_e = y - line_h

    for label, hrs_col, amt_col in rows_earn:
        y_e -= line_h
        c.line(earnings_left, y_e, earnings_right, y_e)
        c.drawString(earnings_left + 5, y_e + 4, label)

        hrs = "" if hrs_col is None else get_safe(row, hrs_col, "")
        amt = get_safe(row, amt_col, 0) if amt_col else ""

        if hrs != "":
            c.drawRightString(earnings_left + 210, y_e + 4, str(hrs))

        if amt != "":
            if label == "Gross Income":
                c.setFont("Helvetica-Bold", 10)
            c.drawRightString(earnings_right - 8, y_e + 4, f"{float(amt):,.2f}")
            if label == "Gross Income":
                c.setFont("Helvetica", 10)

    c.line(earnings_left + 150, earn_bottom, earnings_left + 150, y)
    c.line(earnings_left + 220, earn_bottom, earnings_left + 220, y)

    # ----- DEDUCTIONS BOX -----
    ded_height = line_h * (len(rows_ded) + 1)
    ded_bottom = y - ded_height
    c.rect(ded_left, ded_bottom, ded_right - ded_left, ded_height)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(ded_left + 5, y - line_h + 4, "DEDUCTIONS")
    c.drawString(ded_left + 240, y - line_h + 4, "AMOUNT")
    c.line(ded_left, y - line_h, ded_right, y - line_h)

    c.setFont("Helvetica", 10)
    y_d = y - line_h

    for label, col in rows_ded:
        y_d -= line_h
        c.line(ded_left, y_d, ded_right, y_d)

        c.drawString(ded_left + 5, y_d + 4, label)
        amt = get_safe(row, col, 0)

        if label in ("Total Deductions", "NET PAY"):
            c.setFont("Helvetica-Bold", 10)

        c.drawRightString(ded_right - 8, y_d + 4, f"{float(amt):,.2f}")

        if label in ("Total Deductions", "NET PAY"):
            c.setFont("Helvetica", 10)

    c.line(ded_left + 220, ded_bottom, ded_left + 220, y)

    # ---------- TEXT BLOCK + RECEIVED BY ----------
    footer_top = ded_bottom - 25

    c.setFont("Helvetica", 8)
    wrapped = wrap(footer_text, 90)
    text_y = footer_top
    for line in wrapped:
        c.drawString(ded_left, text_y, line)
        text_y -= 10

    # "Received by" section
    text_y -= 15
    c.setFont("Helvetica", 9)
    c.drawString(ded_left, text_y, "Received by:")

    text_y -= 20
    c.setFont("Helvetica-Bold", 9)
    c.drawString(ded_left + 60, text_y, name)
    underline_y = text_y - 3
    c.line(ded_left + 60, underline_y, ded_left + 260, underline_y)

    c.setFont("Helvetica", 7)
    c.drawString(ded_left + 70, underline_y - 10, "Signature over Printed Name / Date")

    # bottom small texts
    c.setFont("Helvetica", 7)
    c.drawString(left, 30, f"Effectivity Date: {effectivity_date}")
    c.drawRightString(right, 30, document_id)

    c.showPage()
    c.save()

    return file_path
