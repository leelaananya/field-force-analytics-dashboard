# field-force-analytics-dashboard
Field force performance &amp; incentive analytics: SQL validation, Python cleaning, Excel dashboard
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule

raw = pd.read_csv("data/raw_field_data.csv")
clean = pd.read_csv("data/clean_field_data.csv")

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name="Arial", bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(name="Arial", italic=True, size=9, color="595959")
BODY_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", bold=True, size=10)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()

def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ============================================================
# SHEET 1: Raw Data
# ============================================================
ws = wb.active
ws.title = "Raw Data"
ws["A1"] = "Field Force Raw Data (as collected — unvalidated)"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Source: monthly field rep submissions, Jan-Jun 2026. Contains known data quality issues (see Data Validation tab)."
ws["A2"].font = SUBTITLE_FONT

headers = list(raw.columns)
header_row = 4
for j, h in enumerate(headers, start=1):
    ws.cell(row=header_row, column=j, value=h)
style_header_row(ws, header_row, len(headers))

for i, row in raw.iterrows():
    for j, h in enumerate(headers, start=1):
        val = row[h]
        cell = ws.cell(row=header_row + 1 + i, column=j, value=(None if pd.isna(val) else val))
        cell.font = BODY_FONT
        cell.border = BORDER

autosize(ws, [10, 14, 10, 10, 13, 12, 13, 13, 15])
ws.freeze_panes = "A5"
RAW_LAST_ROW = header_row + len(raw)

# ============================================================
# SHEET 2: Clean Data
# ============================================================
ws2 = wb.create_sheet("Clean Data")
ws2["A1"] = "Field Force Clean Data (validated & anomaly-flagged)"
ws2["A1"].font = TITLE_FONT
ws2["A2"] = "Output of the Python automation script (python/data_cleaning_automation.py). Duplicates removed, missing values imputed, out-of-range values corrected."
ws2["A2"].font = SUBTITLE_FONT

headers2 = list(clean.columns)
header_row2 = 4
for j, h in enumerate(headers2, start=1):
    ws2.cell(row=header_row2, column=j, value=h)
style_header_row(ws2, header_row2, len(headers2))

for i, row in clean.iterrows():
    for j, h in enumerate(headers2, start=1):
        val = row[h]
        if isinstance(val, bool) or h == "incentive_anomaly_flag":
            val = "TRUE" if str(val) in ("True", "1", "1.0", "true", "TRUE") else "FALSE"
        cell = ws2.cell(row=header_row2 + 1 + i, column=j, value=(None if pd.isna(val) else val))
        cell.font = BODY_FONT
        cell.border = BORDER

autosize(ws2, [10, 14, 10, 10, 13, 12, 13, 13, 15, 16, 18, 15, 14])
ws2.freeze_panes = "A5"
CLEAN_LAST_ROW = header_row2 + len(clean)

# conditional formatting: highlight anomaly rows
col_letter_anomaly = get_column_letter(headers2.index("incentive_anomaly_flag") + 1)
red_fill = PatternFill("solid", fgColor="FFC7CE")
ws2.conditional_formatting.add(
    f"{col_letter_anomaly}{header_row2+1}:{col_letter_anomaly}{CLEAN_LAST_ROW}",
    CellIsRule(operator="equal", formula=['"TRUE"'], fill=red_fill)
)

# ============================================================
# SHEET 3: Data Validation
# ============================================================
ws3 = wb.create_sheet("Data Validation")
ws3["A1"] = "Data Validation Summary"
ws3["A1"].font = TITLE_FONT
ws3["A2"] = "All figures computed live via formulas against the Raw Data tab."
ws3["A2"].font = SUBTITLE_FONT

raw_rows_range = f"'Raw Data'!A5:A{RAW_LAST_ROW}"
val_rows = [
    ("Total raw rows submitted", f"=COUNTA({raw_rows_range})"),
    ("Missing calls_made values", f"=COUNTBLANK('Raw Data'!F5:F{RAW_LAST_ROW})"),
    ("Missing actual_sales values", f"=COUNTBLANK('Raw Data'!H5:H{RAW_LAST_ROW})"),
    ("Missing incentive_paid values", f"=COUNTBLANK('Raw Data'!I5:I{RAW_LAST_ROW})"),
    ("Duplicate rep-month rows (raw)", f"=SUMPRODUCT((COUNTIFS('Raw Data'!B5:B{RAW_LAST_ROW},'Raw Data'!B5:B{RAW_LAST_ROW},'Raw Data'!D5:D{RAW_LAST_ROW},'Raw Data'!D5:D{RAW_LAST_ROW})>1)*1)-SUMPRODUCT((1/COUNTIFS('Raw Data'!B5:B{RAW_LAST_ROW},'Raw Data'!B5:B{RAW_LAST_ROW},'Raw Data'!D5:D{RAW_LAST_ROW},'Raw Data'!D5:D{RAW_LAST_ROW})))"),
    ("Rows with negative actual_sales", f"=COUNTIF('Raw Data'!H5:H{RAW_LAST_ROW},\"<0\")"),
    ("Rows with negative incentive_paid", f"=COUNTIF('Raw Data'!I5:I{RAW_LAST_ROW},\"<0\")"),
    ("Rows where calls_made > 1.5x calls_planned", f"=SUMPRODUCT(('Raw Data'!F5:F{RAW_LAST_ROW}>'Raw Data'!E5:E{RAW_LAST_ROW}*1.5)*1)"),
    ("Total clean rows (post-validation)", f"=COUNTA('Clean Data'!A5:A{CLEAN_LAST_ROW})"),
]

r = 4
ws3.cell(row=r, column=1, value="Validation Check")
ws3.cell(row=r, column=2, value="Value")
style_header_row(ws3, r, 2)
r += 1
for label, formula in val_rows:
    ws3.cell(row=r, column=1, value=label).font = BODY_FONT
    ws3.cell(row=r, column=1).border = BORDER
    ws3.cell(row=r, column=2, value=formula).font = BODY_FONT
    ws3.cell(row=r, column=2).border = BORDER
    r += 1

ws3.cell(row=r+1, column=1, value="% of raw rows affected by at least one flagged issue").font = BOLD_FONT
ws3.cell(row=r+1, column=2, value=f"=ROUND(100*(B5+B9+B10+B6+B7+B8)/B4,1)&\"%\"").font = BOLD_FONT
# Note: this is an approximate composite indicator combining flag counts (documented assumption)
ws3.cell(row=r+2, column=1, value="Note: composite % is an approximate indicator (some rows may have multiple issues, so this is not a strict unique-row count). See python/data_cleaning_automation.py output for exact per-category counts.").font = SUBTITLE_FONT

autosize(ws3, [45, 20])

# ============================================================
# SHEET 4: Variance Analysis
# ============================================================
ws4 = wb.create_sheet("Variance Analysis")
ws4["A1"] = "Variance Analysis — Actual vs. Target Sales"
ws4["A1"].font = TITLE_FONT
ws4["A2"] = "Computed live via SUMIFS formulas against the Clean Data tab."
ws4["A2"].font = SUBTITLE_FONT

regions = sorted(clean["region"].unique().tolist())
reps = clean[["rep_id", "rep_name", "region"]].drop_duplicates().sort_values("rep_name")

# --- Region-level variance table ---
ws4["A4"] = "Region-Level Variance"
ws4["A4"].font = BOLD_FONT
rheaders = ["Region", "Total Target Sales", "Total Actual Sales", "Variance ($)", "Variance (%)"]
for j, h in enumerate(rheaders, start=1):
    ws4.cell(row=5, column=j, value=h)
style_header_row(ws4, 5, len(rheaders))

clean_region_col = "G"  # region is column C in Clean Data (rep_id,rep_name,region,...) -> check order
# Determine actual column letters from clean headers2
def col_of(name):
    return get_column_letter(headers2.index(name) + 1)

REGION_COL = col_of("region")
TARGET_COL = col_of("target_sales")
ACTUAL_COL = col_of("actual_sales")

start_row = 6
for i, reg in enumerate(regions):
    rr = start_row + i
    ws4.cell(row=rr, column=1, value=reg).font = BODY_FONT
    ws4.cell(row=rr, column=2, value=f"=SUMIFS('Clean Data'!{TARGET_COL}5:{TARGET_COL}{CLEAN_LAST_ROW},'Clean Data'!{REGION_COL}5:{REGION_COL}{CLEAN_LAST_ROW},A{rr})")
    ws4.cell(row=rr, column=3, value=f"=SUMIFS('Clean Data'!{ACTUAL_COL}5:{ACTUAL_COL}{CLEAN_LAST_ROW},'Clean Data'!{REGION_COL}5:{REGION_COL}{CLEAN_LAST_ROW},A{rr})")
    ws4.cell(row=rr, column=4, value=f"=C{rr}-B{rr}")
    ws4.cell(row=rr, column=5, value=f"=ROUND(100*D{rr}/B{rr},2)")
    for c in range(1, 6):
        ws4.cell(row=rr, column=c).font = BODY_FONT
        ws4.cell(row=rr, column=c).border = BORDER
region_table_end = start_row + len(regions) - 1

# --- Rep-level variance table ---
rep_title_row = region_table_end + 3
ws4.cell(row=rep_title_row, column=1, value="Rep-Level Variance").font = BOLD_FONT
rep_header_row = rep_title_row + 1
rep_headers = ["Rep Name", "Region", "Total Target Sales", "Total Actual Sales", "Variance ($)", "Variance (%)"]
for j, h in enumerate(rep_headers, start=1):
    ws4.cell(row=rep_header_row, column=j, value=h)
style_header_row(ws4, rep_header_row, len(rep_headers))

REP_NAME_COL = col_of("rep_name")
rep_start = rep_header_row + 1
for i, (_, rep) in enumerate(reps.iterrows()):
    rr = rep_start + i
    name = rep["rep_name"]
    region = rep["region"]
    ws4.cell(row=rr, column=1, value=name).font = BODY_FONT
    ws4.cell(row=rr, column=2, value=region).font = BODY_FONT
    ws4.cell(row=rr, column=3, value=f"=SUMIFS('Clean Data'!{TARGET_COL}5:{TARGET_COL}{CLEAN_LAST_ROW},'Clean Data'!{REP_NAME_COL}5:{REP_NAME_COL}{CLEAN_LAST_ROW},A{rr})")
    ws4.cell(row=rr, column=4, value=f"=SUMIFS('Clean Data'!{ACTUAL_COL}5:{ACTUAL_COL}{CLEAN_LAST_ROW},'Clean Data'!{REP_NAME_COL}5:{REP_NAME_COL}{CLEAN_LAST_ROW},A{rr})")
    ws4.cell(row=rr, column=5, value=f"=D{rr}-C{rr}")
    ws4.cell(row=rr, column=6, value=f"=ROUND(100*E{rr}/C{rr},2)")
    for c in range(1, 7):
        ws4.cell(row=rr, column=c).border = BORDER
        if c > 1:
            ws4.cell(row=rr, column=c).font = BODY_FONT
rep_table_end = rep_start + len(reps) - 1

autosize(ws4, [16, 14, 18, 18, 14, 13])

# ============================================================
# SHEET 5: Incentive Ops
# ============================================================
ws5 = wb.create_sheet("Incentive Ops")
ws5["A1"] = "Incentive Operations — Anomaly Detection"
ws5["A1"].font = TITLE_FONT
ws5["A2"] = "Expected incentive = 3% of actual sales. Anomaly = paid incentive deviates >15% from expected."
ws5["A2"].font = SUBTITLE_FONT

ws5["A4"] = "Summary"
ws5["A4"].font = BOLD_FONT
INC_PAID_COL = col_of("incentive_paid")
summary_labels = [
    ("Total incentive paid", f"=SUM('Clean Data'!{INC_PAID_COL}5:{INC_PAID_COL}{CLEAN_LAST_ROW})"),
    ("Total expected incentive (3% of actual sales)", f"=ROUND(SUM('Clean Data'!{ACTUAL_COL}5:{ACTUAL_COL}{CLEAN_LAST_ROW})*0.03,2)"),
    ("Records flagged as anomalies (>15% deviation)", f"=COUNTIF('Clean Data'!{col_of('incentive_anomaly_flag')}5:{col_of('incentive_anomaly_flag')}{CLEAN_LAST_ROW},\"TRUE\")"),
    ("% of records flagged as anomalies", f"=ROUND(100*COUNTIF('Clean Data'!{col_of('incentive_anomaly_flag')}5:{col_of('incentive_anomaly_flag')}{CLEAN_LAST_ROW},\"TRUE\")/COUNTA('Clean Data'!A5:A{CLEAN_LAST_ROW}),1)&\"%\""),
]
r = 5
for label, formula in summary_labels:
    ws5.cell(row=r, column=1, value=label).font = BODY_FONT
    ws5.cell(row=r, column=1).border = BORDER
    ws5.cell(row=r, column=2, value=formula).font = BOLD_FONT
    ws5.cell(row=r, column=2).border = BORDER
    r += 1

# Overpaid / underpaid breakdown
r += 1
ws5.cell(row=r, column=1, value="Overpaid vs. Underpaid Breakdown").font = BOLD_FONT
r += 1
breakdown_header_row = r
for j, h in enumerate(["Status", "Count", "Total Incentive Paid"], start=1):
    ws5.cell(row=r, column=j, value=h)
style_header_row(ws5, r, 3)
r += 1
VAR_PCT_COL = col_of("incentive_variance_pct")
statuses = [
    ("Overpaid (>15% above expected)", f">0.15"),
    ("Underpaid (>15% below expected)", f"<-0.15"),
]
overpaid_row = r
ws5.cell(row=r, column=1, value="Overpaid (>15% above expected)").font = BODY_FONT
ws5.cell(row=r, column=2, value=f"=COUNTIF('Clean Data'!{VAR_PCT_COL}5:{VAR_PCT_COL}{CLEAN_LAST_ROW},\">0.15\")").font = BODY_FONT
ws5.cell(row=r, column=3, value=f"=SUMIFS('Clean Data'!{INC_PAID_COL}5:{INC_PAID_COL}{CLEAN_LAST_ROW},'Clean Data'!{VAR_PCT_COL}5:{VAR_PCT_COL}{CLEAN_LAST_ROW},\">0.15\")").font = BODY_FONT
for c in range(1, 4):
    ws5.cell(row=r, column=c).border = BORDER
r += 1
underpaid_row = r
ws5.cell(row=r, column=1, value="Underpaid (>15% below expected)").font = BODY_FONT
ws5.cell(row=r, column=2, value=f"=COUNTIF('Clean Data'!{VAR_PCT_COL}5:{VAR_PCT_COL}{CLEAN_LAST_ROW},\"<-0.15\")").font = BODY_FONT
ws5.cell(row=r, column=3, value=f"=SUMIFS('Clean Data'!{INC_PAID_COL}5:{INC_PAID_COL}{CLEAN_LAST_ROW},'Clean Data'!{VAR_PCT_COL}5:{VAR_PCT_COL}{CLEAN_LAST_ROW},\"<-0.15\")").font = BODY_FONT
for c in range(1, 4):
    ws5.cell(row=r, column=c).border = BORDER
r += 1
within_row = r
ws5.cell(row=r, column=1, value="Within range (\u2264 15% deviation)").font = BODY_FONT
ws5.cell(row=r, column=2, value=f"=COUNTA('Clean Data'!A5:A{CLEAN_LAST_ROW})-B{overpaid_row}-B{underpaid_row}").font = BODY_FONT
ws5.cell(row=r, column=3, value=f"=SUM('Clean Data'!{INC_PAID_COL}5:{INC_PAID_COL}{CLEAN_LAST_ROW})-C{overpaid_row}-C{underpaid_row}").font = BODY_FONT
for c in range(1, 4):
    ws5.cell(row=r, column=c).border = BORDER

autosize(ws5, [42, 16, 22])

# ============================================================
# SHEET 6: HQ Dashboard
# ============================================================
ws6 = wb.create_sheet("HQ Dashboard")
ws6["A1"] = "HQ Dashboard — Field Force Performance Overview"
ws6["A1"].font = TITLE_FONT
ws6["A2"] = "Leadership view: national KPIs, region performance, incentive health."
ws6["A2"].font = SUBTITLE_FONT

kpis = [
    ("Total Actual Sales (all regions)", f"=SUM('Clean Data'!{ACTUAL_COL}5:{ACTUAL_COL}{CLEAN_LAST_ROW})"),
    ("Total Target Sales (all regions)", f"=SUM('Clean Data'!{TARGET_COL}5:{TARGET_COL}{CLEAN_LAST_ROW})"),
    ("Overall Variance (%)", f"=ROUND(100*(B4-B5)/B5,2)&\"%\""),
    ("Total Incentive Paid", f"=SUM('Clean Data'!{INC_PAID_COL}5:{INC_PAID_COL}{CLEAN_LAST_ROW})"),
    ("Incentive Anomalies Flagged", f"=COUNTIF('Clean Data'!{col_of('incentive_anomaly_flag')}5:{col_of('incentive_anomaly_flag')}{CLEAN_LAST_ROW},\"TRUE\")"),
]
r = 4
for label, formula in kpis:
    ws6.cell(row=r, column=1, value=label).font = BOLD_FONT
    ws6.cell(row=r, column=2, value=formula).font = BOLD_FONT
    ws6.cell(row=r, column=1).border = BORDER
    ws6.cell(row=r, column=2).border = BORDER
    r += 1

# Region table for chart source (reference Variance Analysis sheet)
chart_start = r + 2
ws6.cell(row=chart_start-1, column=1, value="Region Variance (%) — Source Data").font = BOLD_FONT
for j, h in enumerate(["Region", "Variance (%)"], start=1):
    ws6.cell(row=chart_start, column=j, value=h)
style_header_row(ws6, chart_start, 2)
for i, reg in enumerate(regions):
    rr = chart_start + 1 + i
    ws6.cell(row=rr, column=1, value=f"='Variance Analysis'!A{start_row+i}")
    ws6.cell(row=rr, column=2, value=f"='Variance Analysis'!E{start_row+i}")
    ws6.cell(row=rr, column=1).border = BORDER
    ws6.cell(row=rr, column=2).border = BORDER
region_chart_data_end = chart_start + len(regions)

bar = BarChart()
bar.title = "Sales Variance by Region (%)"
bar.y_axis.title = "Variance (%)"
bar.x_axis.title = "Region"
data_ref = Reference(ws6, min_col=2, min_row=chart_start, max_row=region_chart_data_end)
cats_ref = Reference(ws6, min_col=1, min_row=chart_start+1, max_row=region_chart_data_end)
bar.add_data(data_ref, titles_from_data=True)
bar.set_categories(cats_ref)
bar.width = 16
bar.height = 9
ws6.add_chart(bar, f"D{chart_start-1}")

# Incentive status pie chart (references Incentive Ops sheet)
pie_start = region_chart_data_end + 3
ws6.cell(row=pie_start-1, column=1, value="Incentive Status Breakdown — Source Data").font = BOLD_FONT
for j, h in enumerate(["Status", "Count"], start=1):
    ws6.cell(row=pie_start, column=j, value=h)
style_header_row(ws6, pie_start, 2)
ws6.cell(row=pie_start+1, column=1, value="Overpaid")
ws6.cell(row=pie_start+1, column=2, value=f"='Incentive Ops'!B{overpaid_row}")
ws6.cell(row=pie_start+2, column=1, value="Underpaid")
ws6.cell(row=pie_start+2, column=2, value=f"='Incentive Ops'!B{underpaid_row}")
ws6.cell(row=pie_start+3, column=1, value="Within range")
ws6.cell(row=pie_start+3, column=2, value=f"='Incentive Ops'!B{within_row}")
for rr in range(pie_start+1, pie_start+4):
    for c in (1, 2):
        ws6.cell(row=rr, column=c).border = BORDER

pie = PieChart()
pie.title = "Incentive Payout Status"
data_ref2 = Reference(ws6, min_col=2, min_row=pie_start, max_row=pie_start+3)
cats_ref2 = Reference(ws6, min_col=1, min_row=pie_start+1, max_row=pie_start+3)
pie.add_data(data_ref2, titles_from_data=True)
pie.set_categories(cats_ref2)
pie.width = 12
pie.height = 9
ws6.add_chart(pie, f"D{pie_start-1}")

autosize(ws6, [34, 16])

# ============================================================
# SHEET 7: Field/Rep Dashboard
# ============================================================
ws7 = wb.create_sheet("Field-Rep Dashboard")
ws7["A1"] = "Field / Rep Dashboard — Individual Scorecard"
ws7["A1"].font = TITLE_FONT
ws7["A2"] = "Rep-level view: ranked by sales variance, with incentive alignment flag."
ws7["A2"].font = SUBTITLE_FONT

header_row7 = 4
scorecard_headers = ["Rank", "Rep Name", "Region", "Target Sales", "Actual Sales", "Variance (%)", "Incentive Anomaly?"]
for j, h in enumerate(scorecard_headers, start=1):
    ws7.cell(row=header_row7, column=j, value=h)
style_header_row(ws7, header_row7, len(scorecard_headers))

# Build ranked list via Python (values), but pull from Variance Analysis via formulas referencing rep table order sorted by variance desc
rep_variance_df = clean.groupby(["rep_name", "region"]).agg(
    target_sales=("target_sales", "sum"),
    actual_sales=("actual_sales", "sum")
).reset_index()
rep_variance_df["variance_pct"] = round(100 * (rep_variance_df["actual_sales"] - rep_variance_df["target_sales"]) / rep_variance_df["target_sales"], 2)
rep_variance_df = rep_variance_df.sort_values("variance_pct", ascending=False).reset_index(drop=True)

anomaly_reps = set(clean.loc[clean["incentive_anomaly_flag"].astype(str).isin(["True", "1", "1.0"]), "rep_name"])

start7 = header_row7 + 1
for i, row in rep_variance_df.iterrows():
    rr = start7 + i
    ws7.cell(row=rr, column=1, value=i + 1).font = BODY_FONT
    ws7.cell(row=rr, column=2, value=row["rep_name"]).font = BODY_FONT
    ws7.cell(row=rr, column=3, value=row["region"]).font = BODY_FONT
    ws7.cell(row=rr, column=4, value=int(row["target_sales"])).font = BODY_FONT
    ws7.cell(row=rr, column=5, value=float(row["actual_sales"])).font = BODY_FONT
    ws7.cell(row=rr, column=6, value=float(row["variance_pct"])).font = BODY_FONT
    ws7.cell(row=rr, column=7, value=("Yes" if row["rep_name"] in anomaly_reps else "No")).font = BODY_FONT
    for c in range(1, 8):
        ws7.cell(row=rr, column=c).border = BORDER
scorecard_end = start7 + len(rep_variance_df) - 1

# conditional formatting on variance column
green_fill = PatternFill("solid", fgColor="C6EFCE")
red_fill2 = PatternFill("solid", fgColor="FFC7CE")
ws7.conditional_formatting.add(
    f"F{start7}:F{scorecard_end}",
    CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=green_fill)
)
ws7.conditional_formatting.add(
    f"F{start7}:F{scorecard_end}",
    CellIsRule(operator="lessThan", formula=["0"], fill=red_fill2)
)
ws7.conditional_formatting.add(
    f"G{start7}:G{scorecard_end}",
    CellIsRule(operator="equal", formula=['"Yes"'], fill=red_fill2)
)

autosize(ws7, [7, 16, 12, 15, 15, 13, 18])
ws7.freeze_panes = "A5"

wb.save("excel/field_force_dashboard.xlsx")
print("Workbook saved.")
