#!/usr/bin/env python3
"""Build Content Creator Revenue Tracker Excel template."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.data_source import NumRef
import os

wb = openpyxl.Workbook()

# Color palette
NAVY = "1B2A4A"
GOLD = "D4A843"
LIGHT_NAVY = "2C3E6B"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F4F7"
MEDIUM_GRAY = "E0E3E8"
DARK_TEXT = "1B2A4A"
GREEN = "27AE60"
RED = "E74C3C"
SOFT_GOLD = "FFF3D6"

header_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
gold_fill = PatternFill(start_color=GOLD, end_color=GOLD, fill_type="solid")
light_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
soft_gold_fill = PatternFill(start_color=SOFT_GOLD, end_color=SOFT_GOLD, fill_type="solid")
kpi_fill_green = PatternFill(start_color="E8F8F0", end_color="E8F8F0", fill_type="solid")
kpi_fill_gold = PatternFill(start_color=SOFT_GOLD, end_color=SOFT_GOLD, fill_type="solid")
thin_border = Border(
    left=Side(style="thin", color=MEDIUM_GRAY),
    right=Side(style="thin", color=MEDIUM_GRAY),
    top=Side(style="thin", color=MEDIUM_GRAY),
    bottom=Side(style="thin", color=MEDIUM_GRAY)
)
header_border = Border(
    left=Side(style="thin", color=GOLD),
    right=Side(style="thin", color=GOLD),
    top=Side(style="thin", color=GOLD),
    bottom=Side(style="thin", color=GOLD)
)

def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border

def style_data_cell(ws, row, col, fmt=None):
    cell = ws.cell(row=row, column=col)
    cell.border = thin_border
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        cell.number_format = fmt
    return cell

def style_input_cell(ws, row, col, fmt=None):
    cell = ws.cell(row=row, column=col)
    cell.border = Border(
        left=Side(style="thin", color=GOLD),
        right=Side(style="thin", color=GOLD),
        top=Side(style="thin", color=GOLD),
        bottom=Side(style="thin", color=GOLD)
    )
    cell.fill = soft_gold_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if fmt:
        cell.number_format = fmt
    return cell

# ============ SHEET 1: SETUP ============
ws_setup = wb.active
ws_setup.title = "Setup"
ws_setup.sheet_properties.tabColor = NAVY

ws_setup.column_dimensions["A"].width = 30
ws_setup.column_dimensions["B"].width = 25
ws_setup.column_dimensions["C"].width = 40

setup_title = ws_setup.cell(row=1, column=1, value="Content Creator Revenue Tracker — Setup")
setup_title.font = Font(name="Calibri", bold=True, color=WHITE, size=14)
setup_title.fill = header_fill
setup_title.alignment = Alignment(horizontal="center")
for c in range(2, 4):
    ws_setup.cell(row=1, column=c).fill = header_fill

# Settings table
settings = [
    ("Parameter", "Value", "Description"),
    ("Currency Symbol", "$", "Change to your local currency (e.g., £, €, ¥)"),
    ("Tax Rate (%)", "25", "Estimated effective tax rate for income estimation"),
    ("Platform 1", "YouTube AdSense", "Primary video platform revenue"),
    ("Platform 2", "Sponsorships", "Brand deal income"),
    ("Platform 3", "Patreon/Memberships", "Recurring membership revenue"),
    ("Platform 4", "Affiliate Income", "Commission from affiliate links"),
    ("Platform 5", "Digital Products", "Ebooks, templates, courses"),
    ("Platform 6", "Merchandise", "Physical product sales"),
    ("Platform 7", "Consulting/Coaching", "1-on-1 or group consulting fees"),
    ("Platform 8", "Speaking Events", "Paid speaking, podcasts, panels"),
    ("Platform 9", "Donations/Tips", "Super Chat, Ko-fi, Buy Me a Coffee"),
    ("Platform 10", "Other Income", "Any other creator income"),
]

for i, (p, v, d) in enumerate(settings):
    row = i + 3
    ws_setup.cell(row=row, column=1, value=p).font = Font(bold=True, color=DARK_TEXT)
    ws_setup.cell(row=row, column=1).fill = light_fill
    ws_setup.cell(row=row, column=1).border = thin_border
    if i == 0:
        style_header_row(ws_setup, row, 3)
    else:
        cell_v = ws_setup.cell(row=row, column=2, value=v)
        cell_v.border = thin_border
        cell_v.alignment = Alignment(horizontal="center")
        if i in [1, 2]:  # currency and tax rate are inputs
            style_input_cell(ws_setup, row, 2)
        ws_setup.cell(row=row, column=3, value=d).border = thin_border
        ws_setup.cell(row=row, column=3).alignment = Alignment(wrap_text=True)

ws_setup.cell(row=16, column=1, value="Revenue Categories").font = Font(bold=True, size=12, color=NAVY)
ws_setup.cell(row=17, column=1, value="This template tracks 10 income streams across all creator platforms.").font = Font(italic=True, size=10, color="666666")

# ============ SHEET 2: INCOME ============
ws_income = wb.create_sheet("Income")
ws_income.sheet_properties.tabColor = GOLD

# Column widths
income_cols = {
    "A": 14, "B": 14, "C": 18, "D": 16, "E": 18, "F": 14,
    "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 25
}
for col, w in income_cols.items():
    ws_income.column_dimensions[col].width = w

# Title row
title = ws_income.cell(row=1, column=1, value="Income Tracker — Log Every Payment")
title.font = Font(name="Calibri", bold=True, color=WHITE, size=14)
title.fill = header_fill
title.alignment = Alignment(horizontal="center")
for c in range(2, 13):
    ws_income.cell(row=1, column=c).fill = header_fill

# Headers (row 2)
income_headers = [
    "Date", "Platform", "Income Type", "Client/Source",
    "Description", "Amount", "Tax Deductible?",
    "Payment Status", "Invoice #", "Quarter",
    "Month", "Notes"
]
for j, h in enumerate(income_headers, 1):
    ws_income.cell(row=2, column=j, value=h)
style_header_row(ws_income, 2, 12)

# Sample data (rows 3-12)
sample_income = [
    ("2026-01-15", "YouTube AdSense", "Ad Revenue", "YouTube", "January AdSense payout", 1250.00, True, "Paid", "", "Q1", "January", ""),
    ("2026-01-20", "Sponsorships", "Brand Deal", "TechCorp Inc", "Product review video sponsorship", 3000.00, True, "Paid", "INV-001", "Q1", "January", "Includes usage rights"),
    ("2026-02-01", "Patreon/Memberships", "Monthly Recurring", "Patreon", "February membership revenue (150 members)", 750.00, False, "Paid", "", "Q1", "February", ""),
    ("2026-02-10", "Affiliate Income", "Commission", "Amazon Associates", "January affiliate commissions", 420.00, False, "Paid", "", "Q1", "February", ""),
    ("2026-02-15", "Digital Products", "Product Sales", "Gumroad", "Spreadsheet template bundle sales", 890.00, True, "Paid", "", "Q1", "February", ""),
    ("2026-03-01", "YouTube AdSense", "Ad Revenue", "YouTube", "February AdSense payout", 1380.00, True, "Paid", "", "Q1", "March", ""),
    ("2026-03-05", "Sponsorships", "Brand Deal", "SkillShare", "Course sponsorship + affiliate", 2500.00, True, "Pending", "INV-002", "Q1", "March", "Net 30 terms"),
    ("2026-03-15", "Merchandise", "Product Sales", "Print-on-demand", "T-shirt and mug sales", 340.00, True, "Paid", "", "Q1", "March", ""),
    ("2026-04-01", "Patreon/Memberships", "Monthly Recurring", "Patreon", "March membership (162 members)", 810.00, False, "Paid", "", "Q2", "April", "12 new members"),
    ("2026-04-10", "Consulting/Coaching", "1-on-1 Session", "Private client", "2-hour strategy session", 500.00, True, "Paid", "INV-003", "Q2", "April", ""),
]

for i, row_data in enumerate(sample_income):
    r = i + 3
    for j, val in enumerate(row_data, 1):
        cell = ws_income.cell(row=r, column=j, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if j == 6:  # Amount column
            cell.number_format = '$#,##0.00'
        if j == 1:  # Date
            cell.number_format = 'YYYY-MM-DD'
        # Alternate row shading
        if i % 2 == 0:
            cell.fill = light_fill

# Add blank rows for user input (rows 13-62)
for r in range(13, 63):
    for c in range(1, 13):
        cell = ws_income.cell(row=r, column=c)
        cell.border = thin_border
        if r % 2 == 0:
            cell.fill = light_fill
        if c == 6:
            cell.number_format = '$#,##0.00'
        if c == 1:
            cell.number_format = 'YYYY-MM-DD'

# Summary formulas at row 64
ws_income.cell(row=64, column=1, value="TOTALS").font = Font(bold=True, color=WHITE, size=11)
ws_income.cell(row=64, column=1).fill = header_fill
for c in range(2, 13):
    ws_income.cell(row=64, column=c).fill = header_fill
    ws_income.cell(row=64, column=c).border = header_border

ws_income.cell(row=64, column=6, value="=SUM(F3:F62)")
ws_income.cell(row=64, column=6).font = Font(bold=True, color=WHITE, size=11)
ws_income.cell(row=64, column=6).number_format = '$#,##0.00'

# ============ SHEET 3: EXPENSES ============
ws_expense = wb.create_sheet("Expenses")
ws_expense.sheet_properties.tabColor = "C0392B"

expense_cols = {
    "A": 14, "B": 14, "C": 20, "D": 18, "E": 22,
    "F": 14, "G": 14, "H": 14, "I": 14, "J": 14, "K": 25
}
for col, w in expense_cols.items():
    ws_expense.column_dimensions[col].width = w

t2 = ws_expense.cell(row=1, column=1, value="Expense Tracker — Track All Business Expenses")
t2.font = Font(name="Calibri", bold=True, color=WHITE, size=14)
t2.fill = header_fill
t2.alignment = Alignment(horizontal="center")
for c in range(2, 12):
    ws_expense.cell(row=1, column=c).fill = header_fill

expense_headers = [
    "Date", "Category", "Sub-Category", "Vendor/Payee",
    "Description", "Amount", "Tax Deductible?",
    "Receipt Saved?", "Quarter", "Month", "Notes"
]
for j, h in enumerate(expense_headers, 1):
    ws_expense.cell(row=2, column=j, value=h)
style_header_row(ws_expense, 2, 11)

# Expense categories reference
expense_categories = [
    "Equipment", "Software/Subscriptions", "Content Production",
    "Marketing/Advertising", "Travel", "Education/Training",
    "Home Office", "Professional Services", "Insurance",
    "Internet/Phone", "Shipping/Supplies", "Platform Fees",
    "Entertainment/Meals", "Legal/Accounting", "Other"
]

sample_expenses = [
    ("2026-01-05", "Software/Subscriptions", "Video Editing", "Adobe Creative Cloud", "Monthly Creative Cloud subscription", 54.99, True, True, "Q1", "January", ""),
    ("2026-01-10", "Equipment", "Camera Gear", "B&H Photo", "New microphone and lighting kit", 350.00, True, True, "Q1", "January", "Depreciable asset"),
    ("2026-01-15", "Home Office", "Internet", "ISP Provider", "Monthly broadband internet", 89.99, True, True, "Q1", "January", "60% business use"),
    ("2026-02-01", "Software/Subscriptions", "Analytics", "TubeBuddy/VidIQ", "YouTube optimization tool", 15.00, True, True, "Q1", "February", ""),
    ("2026-02-05", "Content Production", "Stock Media", "Envato Elements", "Stock footage and music license", 29.00, True, True, "Q1", "February", ""),
    ("2026-02-15", "Marketing/Advertising", "Social Ads", "Meta Ads", "Instagram promotion for new video series", 150.00, True, True, "Q1", "February", ""),
    ("2026-03-01", "Platform Fees", "Payment Processing", "Stripe/PayPal", "Transaction fees on product sales", 45.00, True, True, "Q1", "March", ""),
    ("2026-03-10", "Education/Training", "Online Course", "MasterClass", "Filmmaking course subscription", 180.00, True, True, "Q1", "March", "Annual subscription"),
    ("2026-03-20", "Travel", "Event Travel", "Flight booking", "Flight to CreatorCon conference", 450.00, True, True, "Q1", "March", ""),
    ("2026-04-01", "Software/Subscriptions", "Design Tool", "Canva Pro", "Monthly Canva Pro subscription", 12.99, True, True, "Q2", "April", ""),
]

for i, row_data in enumerate(sample_expenses):
    r = i + 3
    for j, val in enumerate(row_data, 1):
        cell = ws_expense.cell(row=r, column=j, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if j == 6:
            cell.number_format = '$#,##0.00'
        if j == 1:
            cell.number_format = 'YYYY-MM-DD'
        if i % 2 == 0:
            cell.fill = light_fill

for r in range(13, 63):
    for c in range(1, 12):
        cell = ws_expense.cell(row=r, column=c)
        cell.border = thin_border
        if r % 2 == 0:
            cell.fill = light_fill
        if c == 6:
            cell.number_format = '$#,##0.00'
        if c == 1:
            cell.number_format = 'YYYY-MM-DD'

# Total row
ws_expense.cell(row=64, column=1, value="TOTALS").font = Font(bold=True, color=WHITE, size=11)
ws_expense.cell(row=64, column=1).fill = header_fill
for c in range(2, 12):
    ws_expense.cell(row=64, column=c).fill = header_fill
    ws_expense.cell(row=64, column=c).border = header_border
ws_expense.cell(row=64, column=6, value="=SUM(F3:F62)")
ws_expense.cell(row=64, column=6).font = Font(bold=True, color=WHITE, size=11)
ws_expense.cell(row=64, column=6).number_format = '$#,##0.00'

# ============ SHEET 4: DASHBOARD ============
ws_dash = wb.create_sheet("Dashboard")
ws_dash.sheet_properties.tabColor = "27AE60"

dash_cols = {"A": 4, "B": 22, "C": 4, "D": 22, "E": 4, "F": 22, "G": 4, "H": 22, "I": 4, "J": 25}
for col, w in dash_cols.items():
    ws_dash.column_dimensions[col].width = w

# Title
title_cell = ws_dash.cell(row=1, column=1, value="Creator Revenue Dashboard")
title_cell.font = Font(name="Calibri", bold=True, color=WHITE, size=16)
title_cell.fill = header_fill
title_cell.alignment = Alignment(horizontal="center")
for c in range(2, 11):
    ws_dash.cell(row=1, column=c).fill = header_fill

ws_dash.cell(row=2, column=1, value="All metrics auto-calculate from Income and Expenses sheets").font = Font(italic=True, color="666666", size=10)
ws_dash.cell(row=2, column=1).alignment = Alignment(horizontal="center")
for c in range(2, 11):
    ws_dash.cell(row=2, column=c).fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")

# KPI Cards (row 4) - label in col B, value in col C; label in col D, value in col E; etc.
# Using single cells per label and value (no merging for KPI values)
kpis_row1 = [
    ("B4", "Total Revenue", "C4", "=Income!F64", True),
    ("D4", "Total Expenses", "E4", "=Expenses!F64", False),
    ("F4", "Net Profit", "G4", "=C4-E4", True),
    ("H4", "Profit Margin", "I4", "=IF(C4=0,0,G4/C4)", True),
]

for (label_addr, label, val_addr, formula, is_green) in kpis_row1:
    lc = ws_dash[label_addr]
    lc.value = label
    lc.font = Font(bold=True, color=NAVY, size=10)
    lc.fill = light_fill
    lc.alignment = Alignment(horizontal="center")
    lc.border = thin_border

    vc = ws_dash[val_addr]
    vc.value = formula
    vc.font = Font(bold=True, size=16, color=GREEN if is_green else RED)
    vc.fill = kpi_fill_green if is_green else PatternFill(start_color="FDEDEC", end_color="FDEDEC", fill_type="solid")
    vc.alignment = Alignment(horizontal="center")
    vc.border = thin_border
    if "Margin" in label:
        vc.number_format = '0.0%'
    else:
        vc.number_format = '$#,##0.00'

# Second row of KPIs (row 6)
kpis_row2 = [
    ("B6", "Est. Quarterly Tax", "C6", "=ROUND(C6*E6*Setup!B5,0)", True),
    ("D6", "After-Tax Income", "E6", "=C6-E6-G6", True),
    ("F6", "Avg Monthly Income", "G6", "=C6/12", True),
    ("H6", "Income Streams", "I6", "=COUNTA(Income!B3:B62)", True),
]

# Fix: these are interdependent, let me simplify
# Q2: After-Tax = C4-E4 - tax. Let's use direct refs
kpis_row2 = [
    ("B6", "Est. Quarterly Tax", "C6", "=ROUND((C4-E4)*Setup!B5,0)", True),
    ("D6", "After-Tax Income", "E6", "=C4-E4-C6", True),
    ("F6", "Avg Monthly Income", "G6", "=C4/12", True),
    ("H6", "Income Streams", "I6", "=COUNTA(Income!B3:B62)", True),
]

for (label_addr, label, val_addr, formula, is_green) in kpis_row2:
    lc = ws_dash[label_addr]
    lc.value = label
    lc.font = Font(bold=True, color=NAVY, size=10)
    lc.fill = light_fill
    lc.alignment = Alignment(horizontal="center")
    lc.border = thin_border

    vc = ws_dash[val_addr]
    vc.value = formula
    vc.font = Font(bold=True, size=14, color=GOLD)
    vc.fill = kpi_fill_gold
    vc.alignment = Alignment(horizontal="center")
    vc.border = thin_border
    vc.number_format = '$#,##0.00'

# Section: Revenue by Platform (row 9)
ws_dash.cell(row=9, column=2, value="Revenue by Platform").font = Font(bold=True, size=12, color=NAVY)
ws_dash.cell(row=9, column=4, value="Monthly Revenue Trend").font = Font(bold=True, size=12, color=NAVY)
ws_dash.cell(row=9, column=7, value="Expense Breakdown").font = Font(bold=True, size=12, color=NAVY)

# Platform summary table (rows 10-20)
platform_header = ["Platform", "Total Income", "% of Total"]
for j, h in enumerate(platform_header, 2):
    cell = ws_dash.cell(row=10, column=j, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center")
    cell.border = header_border

# Platform totals using SUMIF
for i in range(10):
    r = i + 11
    platform_name = f"Setup!B{i+6}"  # references Setup sheet platform names
    ws_dash.cell(row=r, column=2).value = f"='{platform_name}'"
    ws_dash.cell(row=r, column=2).font = Font(size=10, color=DARK_TEXT)
    ws_dash.cell(row=r, column=2).border = thin_border
    ws_dash.cell(row=r, column=2).fill = light_fill

    # SUMIF formula for each platform
    sumif_cell = ws_dash.cell(row=r, column=3)
    sumif_cell.value = f'=SUMIF(Income!B$3:B$62,B{r},Income!F$3:F$62)'
    sumif_cell.number_format = '$#,##0.00'
    sumif_cell.font = Font(bold=True, size=10)
    sumif_cell.border = thin_border
    sumif_cell.alignment = Alignment(horizontal="center")

    pct_cell = ws_dash.cell(row=r, column=4)
    pct_cell.value = f'=IF(C11=0,0,C{r}/C11)'
    pct_cell.number_format = '0.0%'
    pct_cell.font = Font(size=10)
    pct_cell.border = thin_border
    pct_cell.alignment = Alignment(horizontal="center")

# Tax Schedule Section (row 23)
ws_dash.cell(row=23, column=2, value="Quarterly Tax Schedule").font = Font(bold=True, size=12, color=NAVY)
tax_headers = ["Quarter", "Est. Income", "Est. Tax Due", "Paid?"]
for j, h in enumerate(tax_headers, 2):
    cell = ws_dash.cell(row=24, column=j, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center")
    cell.border = header_border

quarters = [("Q1", 3), ("Q2", 6), ("Q3", 9), ("Q4", 12)]
for i, (q, col) in enumerate(quarters):
    r = i + 25
    ws_dash.cell(row=r, column=2, value=q).font = Font(bold=True)
    ws_dash.cell(row=r, column=2).border = thin_border
    ws_dash.cell(row=r, column=2).fill = light_fill

    ws_dash.cell(row=r, column=3, value=f'=SUMIF(Income!J$3:J$62,B{r},Income!F$3:F$62)')
    ws_dash.cell(row=r, column=3).number_format = '$#,##0.00'
    ws_dash.cell(row=r, column=3).border = thin_border

    ws_dash.cell(row=r, column=4, value=f'=ROUND(C{r}*Setup!B5,0)')
    ws_dash.cell(row=r, column=4).number_format = '$#,##0.00'
    ws_dash.cell(row=r, column=4).border = thin_border
    ws_dash.cell(row=r, column=4).font = Font(bold=True, color=RED)

    ws_dash.cell(row=r, column=5, value="").border = thin_border

# Key Metrics (row 32)
ws_dash.cell(row=32, column=2, value="Key Creator Metrics").font = Font(bold=True, size=12, color=NAVY)

metrics = [
    ("Avg Revenue per Stream", "=IF(I6=0,0,B4/I6)", "$#,##0.00"),
    ("Top Platform (by Revenue)", "=INDEX(B11:B20,MAX(MATCH(MAX(C11:C20),C11:C20,0)),1)", ""),
    ("Total Deductible Expenses", "=SUMIF(Expenses!G$3:G$62,TRUE,Expenses!F$3:F$62)", "$#,##0.00"),
    ("Non-Deductible Expenses", "=SUMIF(Expenses!G$3:G$62,FALSE,Expenses!F$3:F$62)", "$#,##0.00"),
    ("Revenue per Month (Avg)", "=B4/COUNTA(Income!K$3:K$62)", "$#,##0.00"),
]

for i, (label, formula, fmt) in enumerate(metrics):
    r = i + 33
    ws_dash.cell(row=r, column=2, value=label).font = Font(size=10, color=DARK_TEXT)
    ws_dash.cell(row=r, column=2).border = thin_border
    ws_dash.cell(row=r, column=2).fill = light_fill

    val_cell = ws_dash.cell(row=r, column=3)
    val_cell.value = formula
    val_cell.number_format = fmt
    val_cell.font = Font(bold=True, size=11)
    val_cell.border = thin_border
    val_cell.alignment = Alignment(horizontal="center")

# Footer note
ws_dash.cell(row=40, column=2, value="Tip: Update Income and Expenses sheets regularly. Dashboard auto-calculates.").font = Font(italic=True, size=9, color="999999")

# Save
output_dir = os.path.expanduser("~/.hermes/data/content-creator-revenue-tracker")
os.makedirs(output_dir, exist_ok=True)
output_path = os.path.join(output_dir, "content-creator-revenue-tracker.xlsx")
wb.save(output_path)
print(f"Template saved: {output_path}")
