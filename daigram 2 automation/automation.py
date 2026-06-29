import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# =========================
# INPUT / OUTPUT
# =========================
input_file = r"PRP Sample Jun (2).xlsx"
sheet_name = "OneTrust Assessment"
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file = f"Diagram2_Output_{timestamp}.xlsx"

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl")
df.columns = df.columns.str.strip()

required_columns = ["ID", "Organization", "Stage", "Date created", "Tags"]
missing = [c for c in required_columns if c not in df.columns]
if missing:
    raise ValueError(f"Missing required columns: {missing}")

for c in ["Organization", "Stage", "Tags"]:
    df[c] = df[c].astype(str).fillna("").str.strip()

df["Date created"] = pd.to_datetime(df["Date created"], errors="coerce")

# =========================
# FILTERS
# =========================
filtered = df[
    df["Tags"].str.contains(r"(?i)^cyber$", na=False)
    & (df["Date created"].dt.year == 2026)
].copy()

# =========================
# STAGE LOGIC
# =========================
closed_stages = {"Completed", "Under review"}
filtered["Final Stage"] = filtered["Stage"].apply(
    lambda x: "Closed" if str(x).strip() in closed_stages else "Open"
)

# =========================
# ORG MAPPING
# =========================
org_map = {
    "Africa": "AFR",
    "APAC": "APAC",
    "BEES": "GRO",
    "BEES | FINTECH": "GRO",
    "Europe": "Europe",
    "GHQ": "GHQ",
    "South America Zone": "SAZ",
    "North America Zone": "NAZ",
    "Middle America Zone": "MAZ",
}
filtered["Org Display"] = filtered["Organization"].map(lambda x: org_map.get(x, x))

# =========================
# PIVOT TABLE
# =========================
pivot = pd.pivot_table(
    filtered,
    index="Org Display",
    columns="Final Stage",
    values="ID",
    aggfunc="count",
    fill_value=0,
)

for col in ["Open", "Closed"]:
    if col not in pivot.columns:
        pivot[col] = 0

pivot = pivot[["Open", "Closed"]].reset_index()
pivot.columns = ["Row Labels", "Open", "Closed"]
pivot["Grand Total"] = pivot["Open"] + pivot["Closed"]

order = ["AFR", "APAC", "GRO", "Europe", "GHQ", "SAZ", "MAZ", "NAZ"]
pivot["sort_order"] = pivot["Row Labels"].apply(lambda x: order.index(x) if x in order else 999)
pivot = pivot.sort_values(["sort_order", "Row Labels"]).drop(columns="sort_order").reset_index(drop=True)

grand_total_row = pd.DataFrame([{
    "Row Labels": "Grand Total",
    "Open": int(pivot["Open"].sum()),
    "Closed": int(pivot["Closed"].sum()),
    "Grand Total": int(pivot["Grand Total"].sum())
}])
pivot_display = pd.concat([pivot, grand_total_row], ignore_index=True)

# =========================
# KPI DATA
# =========================
baseline_25 = 0.60
q1_26 = 0.32
target_26 = 0.65
closed_total = int(filtered["Final Stage"].eq("Closed").sum())
record_total = int(len(filtered))
q2_26 = round(closed_total / record_total, 2) if record_total else 0

kpi_data = [
    ("Baseline '25", baseline_25, "static"),
    ("Q1 '26", q1_26, "static"),
    ("Q2 '26", q2_26, ""),
    ("Target '26", target_26, "static"),
]

# =========================
# CREATE EXCEL
# =========================
wb = Workbook()
ws = wb.active
ws.title = "Dashboard"

header_fill = PatternFill("solid", fgColor="B7DEE8")
subheader_fill = PatternFill("solid", fgColor="D9EAF7")
grand_fill = PatternFill("solid", fgColor="B7DEE8")
bold = Font(bold=True)
thin = Side(border_style="thin", color="000000")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")

# Filters area
ws["A1"] = "Tags"
ws["B1"] = "Cyber"
ws["A2"] = "Date created"
ws["B2"] = "2026"
for cell in ["A1", "A2"]:
    ws[cell].fill = header_fill
    ws[cell].font = bold
    ws[cell].border = border
for cell in ["B1", "B2"]:
    ws[cell].fill = subheader_fill
    ws[cell].border = border

# Pivot table
start_row = 4
headers = list(pivot_display.columns)
for col_idx, h in enumerate(headers, start=1):
    c = ws.cell(row=start_row, column=col_idx, value=h)
    c.fill = header_fill
    c.font = bold
    c.border = border
    c.alignment = center

for r_idx, row in enumerate(pivot_display.itertuples(index=False), start=start_row + 1):
    for c_idx, value in enumerate(row, start=1):
        c = ws.cell(row=r_idx, column=c_idx, value=value)
        c.border = border
        if c_idx > 1:
            c.alignment = center
    if row[0] == "Grand Total":
        for c_idx in range(1, len(headers) + 1):
            ws.cell(row=r_idx, column=c_idx).fill = grand_fill
            ws.cell(row=r_idx, column=c_idx).font = bold

# KPI table
kpi_start_row = start_row + len(pivot_display) + 5
for idx, text in enumerate(["Metric", "Value", "Remark"], start=1):
    c = ws.cell(row=kpi_start_row, column=idx, value=text)
    c.fill = header_fill
    c.font = bold
    c.border = border
    c.alignment = center

for i, (metric, value, remark) in enumerate(kpi_data, start=kpi_start_row + 1):
    ws.cell(row=i, column=1, value=metric)
    ws.cell(row=i, column=2, value=value)
    ws.cell(row=i, column=3, value=remark)
    for col in [1, 2, 3]:
        ws.cell(row=i, column=col).border = border
    ws.cell(row=i, column=2).number_format = "0%"

note_row = kpi_start_row + len(kpi_data) + 2
ws.cell(row=note_row, column=1, value="Q2 formula")
ws.cell(row=note_row, column=2, value=f"Closed / Grand Total = {closed_total} / {record_total}")
ws.cell(row=note_row, column=1).font = bold

# Chart 1
pivot_chart_end_row = start_row + len(pivot_display) - 1
chart1 = BarChart()
chart1.type = "col"
chart1.grouping = "stacked"
chart1.overlap = 100
chart1.title = "2026 Assessment Status"
chart1.y_axis.title = "Count of ID"
chart1.x_axis.title = "Organization"
chart1.height = 8
chart1.width = 14

data1 = Reference(ws, min_col=2, max_col=3, min_row=start_row, max_row=pivot_chart_end_row - 1)
cats1 = Reference(ws, min_col=1, min_row=start_row + 1, max_row=pivot_chart_end_row - 1)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.dLbls = DataLabelList()
chart1.dLbls.showVal = True
chart1.legend.position = "r"
ws.add_chart(chart1, "F2")

# Chart 2
chart2 = BarChart()
chart2.type = "bar"
chart2.title = "Improve in Supplier Response Time"
chart2.x_axis.title = "Percent"
chart2.y_axis.title = "Metric"
chart2.height = 7
chart2.width = 10
chart2.x_axis.numFmt = "0%"

kpi_data_end = kpi_start_row + len(kpi_data)
data2 = Reference(ws, min_col=2, max_col=2, min_row=kpi_start_row, max_row=kpi_data_end)
cats2 = Reference(ws, min_col=1, min_row=kpi_start_row + 1, max_row=kpi_data_end)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.dLbls = DataLabelList()
chart2.dLbls.showVal = True
chart2.legend = None
ws.add_chart(chart2, "F20")

# Formatting
widths = {'A': 18, 'B': 12, 'C': 12, 'D': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 14}
for col, width in widths.items():
    ws.column_dimensions[col].width = width

ws.freeze_panes = "A5"
wb.save(output_file)
print("SUCCESS")
print(output_file)
