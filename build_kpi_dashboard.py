#!/usr/bin/env python3
"""Build a KPI dashboard workbook (.xlsx) with openpyxl - free, offline.

Creates dist/kpi_dashboard.xlsx with a Data sheet (12 months of demo
metrics) and a Dashboard sheet (KPI cards via cells + native Excel
charts). Open it, then Save As -> .xltx to keep it as a template.

Usage:
    pip install openpyxl
    python3 build_kpi_dashboard.py
"""
import datetime, os, random
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ACCENT = PatternFill("solid", fgColor="0F172A")
CARD = PatternFill("solid", fgColor="1E293B")
WHITE = Font(color="F8FAFC", bold=True, size=14)
BIG = Font(color="22D3EE", bold=True, size=24)


def build():
    random.seed(42)
    wb = Workbook()
    data = wb.active
    data.title = "Data"
    data.append(["Month", "Revenue", "Orders", "Expenses", "NewCustomers"])
    months = []
    y, m = 2025, 12
    for _ in range(12):
        months.append(datetime.date(y, m, 1).strftime("%b %Y"))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    for mon in months:
        data.append([mon, round(random.uniform(180, 320), 1), random.randint(900, 1600),
                     round(random.uniform(120, 210), 1), random.randint(80, 220)])

    dash = wb.create_sheet("Dashboard")
    dash.sheet_view.showGridLines = False
    for col in "ABCDEFGH":
        dash.column_dimensions[col].width = 18
    dash["B2"] = "KPI Dashboard (demo data)"
    dash["B2"].font = Font(color="0F172A", bold=True, size=20)

    kpis = [("Total Revenue (k EGP)", "=SUM(Data!B2:B13)"), ("Total Orders", "=SUM(Data!C2:C13)"),
            ("Avg Expenses (k EGP)", "=AVERAGE(Data!D2:D13)"), ("New Customers", "=SUM(Data!E2:E13)")]
    for i, (label, formula) in enumerate(kpis):
        col = get_column_letter(2 + i * 2)
        for r in (4, 5):
            dash[f"{col}{r}"].fill = CARD
            dash[f"{get_column_letter(3 + i * 2)}{r}"].fill = CARD
        dash[f"{col}4"] = label
        dash[f"{col}4"].font = WHITE
        dash[f"{col}5"] = formula
        dash[f"{col}5"].font = BIG
        dash[f"{col}5"].number_format = "#,##0"
        dash[f"{col}5"].alignment = Alignment(horizontal="center")

    bar = BarChart()
    bar.title = "Revenue vs Expenses"
    bar.add_data(Reference(data, min_col=2, max_col=2, min_row=1, max_row=13), titles_from_data=True)
    bar.add_data(Reference(data, min_col=4, max_col=4, min_row=1, max_row=13), titles_from_data=True)
    bar.set_categories(Reference(data, min_col=1, min_row=2, max_row=13))
    bar.height, bar.width = 9, 22
    dash.add_chart(bar, "B8")

    line = LineChart()
    line.title = "New Customers"
    line.add_data(Reference(data, min_col=5, max_col=5, min_row=1, max_row=13), titles_from_data=True)
    line.set_categories(Reference(data, min_col=1, min_row=2, max_row=13))
    line.height, line.width = 9, 22
    dash.add_chart(line, "B27")

    os.makedirs("dist", exist_ok=True)
    wb.save("dist/kpi_dashboard.xlsx")
    print("Wrote dist/kpi_dashboard.xlsx - open it, then Save As .xltx to keep as a template")


if __name__ == "__main__":
    build()
