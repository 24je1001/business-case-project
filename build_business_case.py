#!/usr/bin/env python3
"""
Automated Business Case & Financial Model Builder
---------------------------------------------------
CLI tool that takes basic financial inputs for a technology implementation
project and generates a client-ready Excel workbook with ROI, NPV, and
payback period calculations plus an embedded trend chart.
"""

import argparse
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# Style constants (industry-standard financial modeling color conventions)
# ---------------------------------------------------------------------------
BLUE_INPUT = Font(name="Arial", size=10, color="0000FF")            # hardcoded inputs
BLACK_FORMULA = Font(name="Arial", size=10, color="000000")         # formulas
HEADER_FONT = Font(name="Arial", size=12, bold=True, color="FFFFFF")
SUBHEADER_FONT = Font(name="Arial", size=10, bold=True, color="000000")
TITLE_FONT = Font(name="Arial", size=16, bold=True, color="1F3864")
LABEL_FONT = Font(name="Arial", size=10, bold=False, color="000000")
KPI_LABEL_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
KPI_VALUE_FONT = Font(name="Arial", size=14, bold=True, color="1F3864")

HEADER_FILL = PatternFill("solid", start_color="1F3864")
SUBHEADER_FILL = PatternFill("solid", start_color="D9E1F2")
KPI_FILL = PatternFill("solid", start_color="1F3864")
INPUT_FILL = PatternFill("solid", start_color="FFF2CC")

THIN = Side(style="thin", color="B7B7B7")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")

CURRENCY_FMT = '$#,##0;($#,##0);"-"'
PCT_FMT = '0.0%;(0.0%);"-"'
YEAR_FMT = '0'


def parse_args():
    p = argparse.ArgumentParser(
        description="Generate a client-ready Excel business case for a technology "
                    "implementation project, complete with ROI, NPV, payback period, "
                    "and an embedded trend chart."
    )
    p.add_argument("--project-name", default="Technology Implementation Project",
                    help="Name of the project (default: %(default)s)")
    p.add_argument("--client-name", default="Client Name",
                    help="Client / company name shown on the cover (default: %(default)s)")
    p.add_argument("--implementation-cost", type=float, default=250000,
                    help="One-time implementation / software cost (default: %(default)s)")
    p.add_argument("--iot-sensor-cost", type=float, default=75000,
                    help="One-time IoT sensor / hardware cost (default: %(default)s)")
    p.add_argument("--annual-savings", type=float, default=120000,
                    help="Expected annual savings once live, Year 1 (default: %(default)s)")
    p.add_argument("--savings-growth", type=float, default=0.03,
                    help="Annual growth rate applied to savings each year, e.g. 0.03 "
                         "for 3%% (default: %(default)s)")
    p.add_argument("--annual-maintenance", type=float, default=15000,
                    help="Recurring annual maintenance / support cost (default: %(default)s)")
    p.add_argument("--discount-rate", type=float, default=0.08,
                    help="Discount rate used for NPV, e.g. 0.08 for 8%% (default: %(default)s)")
    p.add_argument("--years", type=int, default=5,
                    help="Number of projection years, excluding Year 0 (default: %(default)s)")
    p.add_argument("--output", default="business_case.xlsx",
                    help="Output .xlsx file path (default: %(default)s)")
    return p.parse_args()


def style_header_row(ws, row, start_col, end_col, fill, font, height=22):
    ws.row_dimensions[row].height = height
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = CENTER
        cell.border = BOX


def build_workbook(args):
    wb = Workbook()

    # =========================================================
    # SHEET 1: Assumptions
    # =========================================================
    ws_a = wb.active
    ws_a.title = "Assumptions"
    ws_a.sheet_view.showGridLines = False
    ws_a.column_dimensions["A"].width = 32
    ws_a.column_dimensions["B"].width = 20
    ws_a.column_dimensions["C"].width = 45

    ws_a["A1"] = args.project_name
    ws_a["A1"].font = TITLE_FONT
    ws_a["A2"] = f"Prepared for: {args.client_name}"
    ws_a["A2"].font = LABEL_FONT
    ws_a.merge_cells("A1:C1")
    ws_a.merge_cells("A2:C2")

    ws_a["A4"] = "Key Assumptions"
    style_header_row(ws_a, 4, 1, 3, HEADER_FILL, HEADER_FONT)
    ws_a.merge_cells("A4:C4")

    headers = ["Assumption", "Value", "Notes"]
    for i, h in enumerate(headers, start=1):
        cell = ws_a.cell(row=5, column=i, value=h)
        cell.font = SUBHEADER_FONT
        cell.fill = SUBHEADER_FILL
        cell.border = BOX
        cell.alignment = CENTER

    rows = [
        ("Implementation Cost ($)", args.implementation_cost,
         "One-time software / implementation spend, Year 0"),
        ("IoT Sensor Cost ($)", args.iot_sensor_cost,
         "One-time hardware spend, Year 0"),
        ("Expected Annual Savings ($), Year 1", args.annual_savings,
         "Gross run-rate savings once solution is live"),
        ("Annual Savings Growth Rate", args.savings_growth,
         "Applied year-over-year to savings (e.g. efficiency gains)"),
        ("Annual Maintenance / Support Cost ($)", args.annual_maintenance,
         "Recurring cost from Year 1 onward"),
        ("Discount Rate", args.discount_rate,
         "Used to discount future cash flows for NPV"),
        ("Projection Horizon (Years)", args.years,
         "Number of post-implementation years modeled"),
    ]
    start_row = 6
    for i, (label, value, note) in enumerate(rows):
        r = start_row + i
        ws_a.cell(row=r, column=1, value=label).font = LABEL_FONT
        ws_a.cell(row=r, column=1).border = BOX
        vcell = ws_a.cell(row=r, column=2, value=value)
        vcell.font = BLUE_INPUT
        vcell.fill = INPUT_FILL
        vcell.border = BOX
        vcell.alignment = CENTER
        if "Rate" in label:
            vcell.number_format = PCT_FMT
        elif "Years" in label:
            vcell.number_format = YEAR_FMT
        else:
            vcell.number_format = CURRENCY_FMT
        ncell = ws_a.cell(row=r, column=3, value=note)
        ncell.font = Font(name="Arial", size=9, italic=True, color="595959")
        ncell.border = BOX

    IMPL_COST = "Assumptions!$B$6"
    IOT_COST = "Assumptions!$B$7"
    SAVINGS_Y1 = "Assumptions!$B$8"
    SAVINGS_GROWTH = "Assumptions!$B$9"
    MAINTENANCE = "Assumptions!$B$10"
    DISCOUNT_RATE = "Assumptions!$B$11"
    N_YEARS = "Assumptions!$B$12"

    ws_a["A14"] = "Blue = hardcoded input.  Black = formula.  Edit only the yellow cells to run new scenarios."
    ws_a["A14"].font = Font(name="Arial", size=9, italic=True, color="595959")
    ws_a.merge_cells("A14:C14")

    # =========================================================
    # SHEET 2: Financial Model
    # =========================================================
    ws = wb.create_sheet("Financial Model")
    ws.sheet_view.showGridLines = False

    n_years = args.years
    last_col = 1 + n_years + 1  
    ws.column_dimensions["A"].width = 30
    for c in range(2, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15

    ws["A1"] = f"{args.project_name} — Financial Model"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells(f"A1:{get_column_letter(last_col)}1")

    year_row = 3
    ws.cell(row=year_row, column=1, value="Year").font = SUBHEADER_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        c = ws.cell(row=year_row, column=col, value=y)
        c.font = SUBHEADER_FONT
        c.alignment = CENTER
        c.fill = SUBHEADER_FILL
        c.border = BOX
        c.number_format = YEAR_FMT
    ws.cell(row=year_row, column=1).fill = SUBHEADER_FILL
    ws.cell(row=year_row, column=1).border = BOX
    ws.cell(row=year_row, column=1).alignment = CENTER

    r = year_row + 1  
    ws.cell(row=r, column=1, value="Initial Investment ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        cell = ws.cell(row=r, column=col)
        if y == 0:
            cell.value = f"=-({IMPL_COST}+{IOT_COST})"
        else:
            cell.value = 0
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    INVESTMENT_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Annual Savings ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        cell = ws.cell(row=r, column=col)
        if y == 0:
            cell.value = 0
        elif y == 1:
            cell.value = f"={SAVINGS_Y1}"
        else:
            prev_col = get_column_letter(col - 1)
            cell.value = f"={prev_col}{r}*(1+{SAVINGS_GROWTH})"
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    SAVINGS_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Maintenance / Support Cost ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        cell = ws.cell(row=r, column=col)
        cell.value = 0 if y == 0 else f"=-{MAINTENANCE}"
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    MAINT_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Net Cash Flow ($)").font = SUBHEADER_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        col_l = get_column_letter(col)
        cell = ws.cell(row=r, column=col)
        cell.value = f"={col_l}{INVESTMENT_ROW}+{col_l}{SAVINGS_ROW}+{col_l}{MAINT_ROW}"
        cell.font = SUBHEADER_FONT
        cell.number_format = CURRENCY_FMT
        cell.fill = SUBHEADER_FILL
        cell.border = BOX
    NET_CF_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Cumulative Cash Flow ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        col_l = get_column_letter(col)
        cell = ws.cell(row=r, column=col)
        if y == 0:
            cell.value = f"={col_l}{NET_CF_ROW}"
        else:
            prev_col = get_column_letter(col - 1)
            cell.value = f"={prev_col}{r}+{col_l}{NET_CF_ROW}"
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    CUM_CF_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Discount Factor").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        col_l = get_column_letter(col)
        cell = ws.cell(row=r, column=col)
        cell.value = f"=1/(1+{DISCOUNT_RATE})^{col_l}${year_row}"
        cell.font = BLACK_FORMULA
        cell.number_format = '0.000'
        cell.border = BOX
    DISC_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="PV of Net Cash Flow ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        col_l = get_column_letter(col)
        cell = ws.cell(row=r, column=col)
        cell.value = f"={col_l}{NET_CF_ROW}*{col_l}{DISC_ROW}"
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    PV_ROW = r

    r += 1  
    ws.cell(row=r, column=1, value="Cumulative PV ($)").font = LABEL_FONT
    for y in range(0, n_years + 1):
        col = 2 + y
        col_l = get_column_letter(col)
        cell = ws.cell(row=r, column=col)
        if y == 0:
            cell.value = f"={col_l}{PV_ROW}"
        else:
            prev_col = get_column_letter(col - 1)
            cell.value = f"={prev_col}{r}+{col_l}{PV_ROW}"
        cell.font = BLACK_FORMULA
        cell.number_format = CURRENCY_FMT
        cell.border = BOX
    CUM_PV_ROW = r

    first_yr1_col = get_column_letter(3)          
    last_yr_col = get_column_letter(2 + n_years)   
    year0_col = get_column_letter(2)

    # KPI summary block
    kpi_row = r + 3
    ws.cell(row=kpi_row, column=1, value="Key Metrics").font = HEADER_FONT
    style_header_row(ws, kpi_row, 1, min(5, last_col), HEADER_FILL, HEADER_FONT)
    ws.merge_cells(start_row=kpi_row, start_column=1, end_row=kpi_row, end_column=min(5, last_col))

    label_row = kpi_row + 1
    value_row = kpi_row + 2

    kpis = [
        ("Net Present Value (NPV)",
         f"=NPV({DISCOUNT_RATE},{first_yr1_col}{NET_CF_ROW}:{last_yr_col}{NET_CF_ROW})+{year0_col}{NET_CF_ROW}",
         CURRENCY_FMT),
        ("ROI (Total)",
         f"=SUM({first_yr1_col}{NET_CF_ROW}:{last_yr_col}{NET_CF_ROW})/-{year0_col}{NET_CF_ROW}",
         PCT_FMT),
        ("Payback Period (Years)",
         (f'=IFERROR('
          f'MATCH(TRUE,INDEX({year0_col}{CUM_CF_ROW}:{last_yr_col}{CUM_CF_ROW}>0,0))-2'
          f'+(-INDEX({year0_col}{CUM_CF_ROW}:{last_yr_col}{CUM_CF_ROW},'
          f'MATCH(TRUE,INDEX({year0_col}{CUM_CF_ROW}:{last_yr_col}{CUM_CF_ROW}>0,0))-1))'
          f'/INDEX({year0_col}{NET_CF_ROW}:{last_yr_col}{NET_CF_ROW},'
          f'MATCH(TRUE,INDEX({year0_col}{CUM_CF_ROW}:{last_yr_col}{CUM_CF_ROW}>0,0))),'
          f'"Not within horizon")'),
         '0.00" yrs"'),
        ("Total Net Savings (undiscounted)",
         f"=SUM({first_yr1_col}{SAVINGS_ROW}:{last_yr_col}{SAVINGS_ROW})"
         f"+SUM({first_yr1_col}{MAINT_ROW}:{last_yr_col}{MAINT_ROW})",
         CURRENCY_FMT),
    ]

    for i, (label, formula, fmt) in enumerate(kpis):
        col = 1 + i
        lcell = ws.cell(row=label_row, column=col, value=label)
        lcell.font = KPI_LABEL_FONT
        lcell.fill = KPI_FILL
        lcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        lcell.border = BOX
        ws.row_dimensions[label_row].height = 30

        vcell = ws.cell(row=value_row, column=col, value=formula)
        vcell.font = KPI_VALUE_FONT
        vcell.alignment = CENTER
        vcell.border = BOX
        vcell.number_format = fmt
        ws.row_dimensions[value_row].height = 26

    # =========================================================
    # Embedded trend chart with forced scales
    # =========================================================
    chart = LineChart()
    chart.title = "Cash Flow Trend"
    chart.style = 2
    chart.y_axis.title = "USD ($)"
    chart.x_axis.title = "Year"
    chart.height = 12
    chart.width = 22

    # CRITICAL FIX: Explicitly force display of axes scale labels 
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.numFmt = "$#,##0"  # Set scaling number formatting on the axis

    cats = Reference(ws, min_col=2, max_col=last_col, min_row=year_row, max_row=year_row)

    data_net = Reference(ws, min_col=1, max_col=last_col, min_row=NET_CF_ROW, max_row=NET_CF_ROW)
    data_cum = Reference(ws, min_col=1, max_col=last_col, min_row=CUM_CF_ROW, max_row=CUM_CF_ROW)
    data_cumpv = Reference(ws, min_col=1, max_col=last_col, min_row=CUM_PV_ROW, max_row=CUM_PV_ROW)

    chart.add_data(data_net, titles_from_data=True, from_rows=True)
    chart.add_data(data_cum, titles_from_data=True, from_rows=True)
    chart.add_data(data_cumpv, titles_from_data=True, from_rows=True)
    chart.set_categories(cats)

    for series in chart.series:
        series.marker = Marker(symbol="circle", size=6)
        series.smooth = False

    anchor_row = value_row + 3
    ws.add_chart(chart, f"A{anchor_row}")
    ws.freeze_panes = "B4"

    # =========================================================
    # SHEET 3: Notes
    # =========================================================
    ws_n = wb.create_sheet("Notes")
    ws_n.sheet_view.showGridLines = False
    ws_n.column_dimensions["A"].width = 100
    ws_n["A1"] = "How this model works"
    ws_n["A1"].font = TITLE_FONT
    notes = [
        "1. All inputs live on the 'Assumptions' tab in yellow/blue cells — change any of them.",
        "2. Year 0 captures one-time implementation and hardware costs.",
        "3. Remember to click 'Enable Editing' at the top of Excel to load values and chart dimensions."
    ]
    for i, note in enumerate(notes):
        cell = ws_n.cell(row=3 + i, column=1, value=note)
        cell.font = Font(name="Arial", size=10)

    return wb


def main():
    args = parse_args()
    wb = build_workbook(args)
    wb.save(args.output)
    print(f"Business case workbook written to: {args.output}")


if __name__ == "__main__":
    main()
    