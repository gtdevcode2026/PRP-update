import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# =========================
# CONFIG
# =========================
input_file = "PRP Sample Jun (2).xlsx"
output_file = "PRP_Final_Output3.xlsx"
sheet_name = "TPRM Web-Portal Export"

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(input_file, sheet_name=sheet_name, engine="openpyxl")

df.columns = df.columns.str.strip().str.lower()

print("Available columns:", df.columns.tolist())

# =========================
# PIVOT TABLE
# =========================
pivot = pd.pivot_table(
    df,
    index="zone_assessing",
    columns="assessment_status",
    values="id",
    aggfunc="count",
    fill_value=0,
    margins=True,
    margins_name="Grand Total"
)

for col in ["ACTIVE", "Active", "Deprioritized", "Duplicate"]:
    if col not in pivot.columns:
        pivot[col] = 0

pivot["Active_Total"] = pivot.get("ACTIVE", 0) + pivot.get("Active", 0)

# Flatten MultiIndex columns (pivot_table produces level "id" on top)
pivot.columns = [
    c[-1] if isinstance(c, tuple) else c
    for c in pivot.columns
]
pivot.index.name = "Zone"
pivot_out = pivot.reset_index()

# =========================
# SUMMARY (for chart 1)
# =========================
grand = pivot.loc["Grand Total"]
summary_df = pd.DataFrame({
    "Status": ["Active", "Deprioritized", "Duplicate"],
    "Count": [
        int(grand.get("Active_Total", 0)),
        int(grand.get("Deprioritized", 0)),
        int(grand.get("Duplicate", 0)),
    ]
})

# =========================
# ZONE ACTIVE TABLE (for chart 2)
# =========================
zone_active = pivot_out[pivot_out["Zone"] != "Grand Total"][["Zone", "Active_Total"]].copy()
zone_active.columns = ["Zone", "Active"]
zone_active = zone_active.reset_index(drop=True)

# =========================
# WRITE — one table per sheet
# =========================
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    pivot_out.to_excel(writer, sheet_name="Pivot", index=False)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    zone_active.to_excel(writer, sheet_name="Active by Zone", index=False)

# =========================
# ADD CHARTS via openpyxl
# =========================
wb = load_workbook(output_file)

# --- Chart 1: Status Overview (on Summary sheet) ---
ws_summary = wb["Summary"]
n_summary = len(summary_df)

chart1 = BarChart()
chart1.title = "Assessment Status Overview"
chart1.y_axis.title = "Count"
chart1.x_axis.title = "Assessment Status"

data1 = Reference(ws_summary, min_col=2, min_row=1, max_row=1 + n_summary)
cats1 = Reference(ws_summary, min_col=1, min_row=2, max_row=1 + n_summary)

chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.dLbls = DataLabelList()
chart1.dLbls.showVal = True

ws_summary.add_chart(chart1, "D2")

# --- Chart 2: Active by Zone (on Active by Zone sheet) ---
ws_zone = wb["Active by Zone"]
n_zones = len(zone_active)

chart2 = BarChart()
chart2.title = "Active by Zone"
chart2.y_axis.title = "Count"
chart2.x_axis.title = "Zone"

data2 = Reference(ws_zone, min_col=2, min_row=1, max_row=1 + n_zones)
cats2 = Reference(ws_zone, min_col=1, min_row=2, max_row=1 + n_zones)

chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.dLbls = DataLabelList()
chart2.dLbls.showVal = True

ws_zone.add_chart(chart2, "D2")

wb.save(output_file)

print("Done. Output:", output_file)
print("Summary:")
print(summary_df.to_string(index=False))
print("Active by Zone:")
print(zone_active.to_string(index=False))
