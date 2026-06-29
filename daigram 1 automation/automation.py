import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# =========================
# INPUT
# =========================
input_file = "PRP Sample Jun (2).xlsx"

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(input_file, sheet_name="TPRM Web-Portal Export", engine="openpyxl")
print("\nColumns in file:")
for col in df.columns:
    print(repr(col))

# Clean columns
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
# =========================
# FILTER DATA (DO THIS FIRST)
# =========================
df["request_date"] = pd.to_datetime(df["request_date"], errors="coerce")

filtered_df = df[
    (df["request_date"].dt.year == 2026) &
    (df["category"].astype(str).str.strip().str.upper() == "TECHNOLOGY")
].copy()

print("Rows after filter:", len(filtered_df))  # should be 72


# =========================
# CREATE PIVOT (USE filtered_df)
# =========================
pivot = (
    filtered_df.groupby("supplier_zone")["id"]
    .count()
    .reset_index(name="Supplier Added by Zone")
)

pivot = pivot.rename(columns={"supplier_zone": "Zone"})

print("\nPivot result:")
print(pivot)

# =========================
# STATIC DATA
# =========================
tier1_data = {
    "Zone": ["NAZ", "AFR", "GHQ", "EUR", "APAC", "SAZ", "MAZ"],
    "Tier-1 Supplier": [123, 115, 79, 35, 17, 13, 13]
}

tier_df = pd.DataFrame(tier1_data)

# =========================
# MERGE
# =========================
final_df = tier_df.merge(pivot, on="Zone", how="left")
final_df["Supplier Added by Zone"] = final_df["Supplier Added by Zone"].fillna(0).astype(int)

# =========================
# ADD TOTAL ROW
# =========================
total_row = pd.DataFrame({
    "Zone": ["Total"],
    "Tier-1 Supplier": [final_df["Tier-1 Supplier"].sum()],
    "Supplier Added by Zone": [final_df["Supplier Added by Zone"].sum()]
})

final_df = pd.concat([final_df, total_row], ignore_index=True)

# =========================
# ✅ CREATE PROFESSIONAL CHART (MATCH IMAGE)
# =========================
chart_df = final_df[final_df["Zone"] != "Total"]

plt.figure(figsize=(10, 6), facecolor="black")

# bars
bar1 = plt.bar(chart_df["Zone"], chart_df["Tier-1 Supplier"],
               color="#1f77b4", label="Tier-1 Supplier")
bar2 = plt.bar(chart_df["Zone"],
               chart_df["Supplier Added by Zone"],
               bottom=chart_df["Tier-1 Supplier"],
               color="#ff7f0e", label="Supplier Added by Zone")

# labels inside bars
for i, v in enumerate(chart_df["Tier-1 Supplier"]):
    plt.text(i, v/2, str(v), ha='center', color='white', fontsize=10)

for i, v in enumerate(chart_df["Supplier Added by Zone"]):
    if v > 0:
        plt.text(i, chart_df["Tier-1 Supplier"][i] + v/2,
                 str(v), ha='center', color='white', fontsize=10)

# styling
plt.title(f"({final_df['Tier-1 Supplier'][:-1].sum()}) Zone wise Tier 1 Suppliers",
          color="white", fontsize=14, weight="bold")

plt.xticks(color="white")
plt.yticks(color="white")

plt.legend(facecolor="black", labelcolor="white")

plt.gca().set_facecolor("black")

# border color
for spine in plt.gca().spines.values():
    spine.set_edgecolor("white")

plt.tight_layout()

chart_file = "chart.png"
plt.savefig(chart_file, facecolor="black")
plt.close()

# =========================
# EXPORT TO EXCEL
# =========================
output_file = "PRP_Output.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Final Table", index=False)

# =========================
# INSERT CHART INTO EXCEL
# =========================
wb = load_workbook(output_file)
ws = wb["Final Table"]

img = Image(chart_file)
ws.add_image(img, "E2")

wb.save(output_file)

print("✅ Output correct with 2026 filter + styled chart inside Excel")