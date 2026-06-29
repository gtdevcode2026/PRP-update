import pandas as pd
import re

# =========================================================
# INPUT / OUTPUT FILE PATH
# =========================================================
input_file = "PRP Sample Jun (2).xlsx"
output_file = "Risk_Output.xlsx"
source_sheet = "OneTrust - Risk Export"

# =========================================================
# LOAD DATA
# =========================================================
df = pd.read_excel(input_file, sheet_name=source_sheet, engine="openpyxl")

# Clean column names including hidden line breaks like Stage\r or Aging\r
df.columns = (
    df.columns.astype(str)
    .str.replace("\r", "", regex=False)
    .str.replace("\n", "", regex=False)
    .str.strip()
)

# =========================================================
# FIND REQUIRED COLUMNS SAFELY
# =========================================================
def find_column(possible_names, columns):
    lookup = {str(col).strip().lower(): col for col in columns}

    for name in possible_names:
        if name.lower() in lookup:
            return lookup[name.lower()]

    return None


organization_col = find_column(["Organization"], df.columns)
id_col = find_column(["ID"], df.columns)
stage_col = find_column(["Stage"], df.columns)
aging_col = find_column(["Aging", "Ageing"], df.columns)

missing = []

if organization_col is None:
    missing.append("Organization")

if id_col is None:
    missing.append("ID")

if stage_col is None:
    missing.append("Stage")

if aging_col is None:
    missing.append("Aging or Ageing")

if missing:
    raise Exception(
        f"Missing required columns: {missing}. Available columns: {list(df.columns)}"
    )

# =========================================================
# CLEAN DATA
# =========================================================
df[organization_col] = df[organization_col].fillna("").astype(str).str.strip()
df[stage_col] = df[stage_col].fillna("").astype(str).str.strip()

# Normalize Stage for filter
df["Stage_Clean"] = df[stage_col].str.lower().str.strip()

# =========================================================
# AGING DAYS LOGIC
# Aging column is used as number of days:
# Aging > 90  = Overdue
# Aging <= 90 = Open
# =========================================================
def extract_number(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    match = re.search(r"\d+(\.\d+)?", text)

    if match:
        return float(match.group())

    return None


df["Aging_Days"] = df[aging_col].apply(extract_number)

# Blank/non-numeric Aging is treated as 0 days, therefore Open
df["Aging_Days"] = df["Aging_Days"].fillna(0)

# =========================================================
# APPLY STAGE FILTER
# Stage filter = Evaluation, Identified, Treatment
# =========================================================
selected_stage_values = ["evaluation", "identified", "treatment"]

filtered_df = df[df["Stage_Clean"].isin(selected_stage_values)].copy()
filtered_df = filtered_df[filtered_df[organization_col] != ""]

# =========================================================
# CLASSIFY OPEN VS OVERDUE
# =========================================================
filtered_df["Risk_Status"] = filtered_df["Aging_Days"].apply(
    lambda x: "Overdue" if x > 90 else "Open"
)

# =========================================================
# CREATE OPEN VS OVERDUE PIVOT
# Rows    = Organization
# Columns = Risk_Status
# Values  = Count of ID
# =========================================================
if len(filtered_df) > 0:
    open_overdue_pivot = pd.pivot_table(
        filtered_df,
        index=organization_col,
        columns="Risk_Status",
        values=id_col,
        aggfunc="count",
        fill_value=0,
        margins=True,
        margins_name="Grand Total"
    ).reset_index()
else:
    open_overdue_pivot = pd.DataFrame(
        columns=[organization_col, "Open", "Overdue", "Grand Total"]
    )

open_overdue_pivot.columns.name = None
open_overdue_pivot.columns = [str(col).strip() for col in open_overdue_pivot.columns]

# Rename first column to Organization for consistent output
if organization_col in open_overdue_pivot.columns:
    open_overdue_pivot = open_overdue_pivot.rename(
        columns={organization_col: "Organization"}
    )

# Ensure all expected columns exist
if "Open" not in open_overdue_pivot.columns:
    open_overdue_pivot["Open"] = 0

if "Overdue" not in open_overdue_pivot.columns:
    open_overdue_pivot["Overdue"] = 0

if "Grand Total" not in open_overdue_pivot.columns:
    open_overdue_pivot["Grand Total"] = (
        open_overdue_pivot["Open"] + open_overdue_pivot["Overdue"]
    )

open_overdue_pivot = open_overdue_pivot[
    ["Organization", "Open", "Overdue", "Grand Total"]
]

# Convert numbers safely
for col in ["Open", "Overdue", "Grand Total"]:
    open_overdue_pivot[col] = pd.to_numeric(
        open_overdue_pivot[col],
        errors="coerce"
    ).fillna(0).astype(int)

# =========================================================
# MAP ORGANIZATION TO ZONES
# =========================================================
zone_map = {
    "Africa": "AFR",
    "APAC": "APAC",
    "BEES": "GRO",
    "BEES | FINTECH": "GRO",
    "Europe": "EUR",
    "GHQ": "GHQ",
    "Middle America Zone": "MAZ",
    "North America Zone": "NAZ",
    "South America Zone": "SAZ"
}

zone_source = open_overdue_pivot[
    open_overdue_pivot["Organization"] != "Grand Total"
].copy()

zone_source["Zones"] = zone_source["Organization"].map(zone_map)

zones_open_overdue = zone_source[["Zones", "Open", "Overdue"]].copy()
zones_open_overdue = zones_open_overdue.dropna(subset=["Zones"])

zones_open_overdue["Open"] = pd.to_numeric(
    zones_open_overdue["Open"],
    errors="coerce"
).fillna(0).astype(int)

zones_open_overdue["Overdue"] = pd.to_numeric(
    zones_open_overdue["Overdue"],
    errors="coerce"
).fillna(0).astype(int)

# Group duplicate mapped zones, e.g. BEES + BEES | FINTECH = GRO
zones_open_overdue = (
    zones_open_overdue
    .groupby("Zones", as_index=False)[["Open", "Overdue"]]
    .sum()
)

# Sort zones as required
zone_order = ["AFR", "APAC", "GRO", "EUR", "GHQ", "MAZ", "NAZ", "SAZ"]

zones_open_overdue["Sort_Order"] = zones_open_overdue["Zones"].apply(
    lambda x: zone_order.index(x) if x in zone_order else 999
)

zones_open_overdue = (
    zones_open_overdue
    .sort_values("Sort_Order")
    .drop(columns=["Sort_Order"])
)

# =========================================================
# ZONE SUMMARY TABLE
# Total Risks = Evaluation, Identified, Treatment, Monitoring
# Open Risks  = Evaluation, Identified, Treatment
# =========================================================
total_stage_values = ["evaluation", "identified", "treatment", "monitoring"]
open_stage_values = ["evaluation", "identified", "treatment"]

total_df = df[df["Stage_Clean"].isin(total_stage_values)].copy()
open_df = df[df["Stage_Clean"].isin(open_stage_values)].copy()

total_counts = (
    total_df
    .groupby(organization_col, as_index=False)[id_col]
    .count()
    .rename(columns={id_col: "Total Risks"})
)

open_counts = (
    open_df
    .groupby(organization_col, as_index=False)[id_col]
    .count()
    .rename(columns={id_col: "Open Risks"})
)

zones_df = pd.merge(
    total_counts,
    open_counts,
    on=organization_col,
    how="left"
)

zones_df["Open Risks"] = zones_df["Open Risks"].fillna(0)
zones_df["Total Risks"] = zones_df["Total Risks"].astype(int)
zones_df["Open Risks"] = zones_df["Open Risks"].astype(int)

zones_df = zones_df.rename(columns={organization_col: "Zones"})
zones_df = zones_df[zones_df["Zones"] != ""]

# =========================================================
# WRITE OUTPUT EXCEL
# =========================================================
with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:

    open_overdue_pivot.to_excel(
        writer,
        sheet_name="Open Overdue Pivot",
        index=False
    )

    zones_open_overdue.to_excel(
        writer,
        sheet_name="Open vs Overdue",
        index=False
    )

    zones_df.to_excel(
        writer,
        sheet_name="Zone Summary",
        index=False
    )

    workbook = writer.book

    # =====================================================
    # FORMATS
    # =====================================================
    header_format = workbook.add_format({
        "bold": True,
        "font_color": "black",
        "bg_color": "#BFEFFF",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    cell_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    total_format = workbook.add_format({
        "bold": True,
        "font_color": "black",
        "bg_color": "#BFEFFF",
        "border": 1,
        "align": "center",
        "valign": "vcenter"
    })

    red_number_format = workbook.add_format({
        "font_color": "red",
        "bold": True,
        "align": "center",
        "valign": "vcenter"
    })

    # =====================================================
    # FORMAT OPEN OVERDUE PIVOT SHEET
    # =====================================================
    ws_pivot = writer.sheets["Open Overdue Pivot"]

    for col_num, col_name in enumerate(open_overdue_pivot.columns):
        ws_pivot.write(0, col_num, col_name, header_format)

    ws_pivot.set_column("A:A", 30, cell_format)
    ws_pivot.set_column("B:D", 15, cell_format)

    if len(open_overdue_pivot) > 0:
        grand_total_excel_row = len(open_overdue_pivot)

        for col_num, col_name in enumerate(open_overdue_pivot.columns):
            ws_pivot.write(
                grand_total_excel_row,
                col_num,
                open_overdue_pivot.iloc[-1, col_num],
                total_format
            )

    # =====================================================
    # FORMAT OPEN VS OVERDUE SHEET
    # =====================================================
    ws_open_overdue = writer.sheets["Open vs Overdue"]

    for col_num, col_name in enumerate(zones_open_overdue.columns):
        ws_open_overdue.write(0, col_num, col_name, header_format)

    ws_open_overdue.set_column("A:A", 15, cell_format)
    ws_open_overdue.set_column("B:C", 15, cell_format)

    overdue_total = (
        int(zones_open_overdue["Overdue"].sum())
        if len(zones_open_overdue) > 0
        else 0
    )

    total_row_position = len(zones_open_overdue) + 2
    ws_open_overdue.write(total_row_position, 2, overdue_total, red_number_format)

    # =====================================================
    # FORMAT ZONE SUMMARY SHEET
    # =====================================================
    ws_zone = writer.sheets["Zone Summary"]

    for col_num, col_name in enumerate(zones_df.columns):
        ws_zone.write(0, col_num, col_name, header_format)

    ws_zone.set_column("A:A", 30, cell_format)
    ws_zone.set_column("B:C", 15, cell_format)

    # =====================================================
    # CHART 1: OPEN VS OVERDUE
    # =====================================================
    if len(zones_open_overdue) > 0:

        chart = workbook.add_chart({
            "type": "column",
            "subtype": "stacked"
        })

        max_row = len(zones_open_overdue) + 1

        categories = f"='Open vs Overdue'!$A$2:$A${max_row}"
        open_values = f"='Open vs Overdue'!$B$2:$B${max_row}"
        overdue_values = f"='Open vs Overdue'!$C$2:$C${max_row}"

        chart.add_series({
            "name": "Open",
            "categories": categories,
            "values": open_values,
            "fill": {"color": "#156082"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart.add_series({
            "name": "Overdue",
            "categories": categories,
            "values": overdue_values,
            "fill": {"color": "#C00000"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart.set_title({
            "name": "Open vs Overdue Risks",
            "name_font": {
                "color": "white",
                "bold": True,
                "size": 14
            }
        })

        chart.set_x_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"}
        })

        chart.set_y_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"},
            "major_gridlines": {
                "visible": True,
                "line": {"color": "#444444"}
            }
        })

        chart.set_legend({
            "position": "bottom",
            "font": {"color": "white"}
        })

        chart.set_chartarea({
            "fill": {"color": "black"}
        })

        chart.set_plotarea({
            "fill": {"color": "black"}
        })

        chart.set_size({
            "width": 720,
            "height": 420
        })

        ws_open_overdue.insert_chart("E4", chart)

    # =====================================================
    # CHART 2: ZONE SUMMARY
    # =====================================================
    if len(zones_df) > 0:

        chart2 = workbook.add_chart({
            "type": "column",
            "subtype": "stacked"
        })

        max_row2 = len(zones_df) + 1

        categories2 = f"='Zone Summary'!$A$2:$A${max_row2}"
        total_values = f"='Zone Summary'!$B$2:$B${max_row2}"
        open_risk_values = f"='Zone Summary'!$C$2:$C${max_row2}"

        chart2.add_series({
            "name": "Total Risks",
            "categories": categories2,
            "values": total_values,
            "fill": {"color": "#156082"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart2.add_series({
            "name": "Open Risks",
            "categories": categories2,
            "values": open_risk_values,
            "fill": {"color": "#F26C23"},
            "border": {"none": True},
            "data_labels": {
                "value": True,
                "font": {"color": "white"}
            }
        })

        chart2.set_title({
            "name": "Zone wise Risks",
            "name_font": {
                "color": "white",
                "bold": True,
                "size": 14
            }
        })

        chart2.set_x_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"}
        })

        chart2.set_y_axis({
            "num_font": {"color": "white"},
            "line": {"color": "white"},
            "major_gridlines": {
                "visible": True,
                "line": {"color": "#444444"}
            }
        })

        chart2.set_legend({
            "position": "bottom",
            "font": {"color": "white"}
        })

        chart2.set_chartarea({
            "fill": {"color": "black"}
        })

        chart2.set_plotarea({
            "fill": {"color": "black"}
        })

        chart2.set_size({
            "width": 720,
            "height": 420
        })

        ws_zone.insert_chart("E4", chart2)

print("Excel file created successfully:", output_file)
print("Aging column used:", aging_col)
print("Open vs Overdue table:")
print(zones_open_overdue)
print(
    "Total Overdue risks where Aging > 90:",
    int(zones_open_overdue["Overdue"].sum()) if len(zones_open_overdue) > 0 else 0
)
