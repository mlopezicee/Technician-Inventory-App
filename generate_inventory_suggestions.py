import pandas as pd

# Load input files
part_file = "rs_PartHistory_rpt.xlsx"
inventory_file = "rs_JdeInventory.xlsx"

# Read part history
part_df = pd.read_excel(part_file, sheet_name=0, engine="openpyxl", header=2)
part_df["Part #"] = part_df["Part #"].astype(str).str.strip()
part_df["Technician"] = part_df["Technician"].astype(str).str.strip()

# Aggregate usage
usage_summary = part_df.groupby(["Technician", "Part #", "Part Description"], as_index=False)["Qty"].sum()
usage_summary.rename(columns={"Qty": "Quantity Used"}, inplace=True)

# Read inventory
inv_df = pd.read_excel(inventory_file, sheet_name=0, engine="openpyxl", header=2)
inv_df["Part #"] = inv_df["Part #"].astype(str).str.strip()
inv_df["Technician"] = inv_df["Technician"].astype(str).str.strip()

# Aggregate inventory
inventory_summary = inv_df.groupby(["Technician", "Part #", "Part Description"], as_index=False)["QoH"].sum()

# Merge usage and inventory
merged = pd.merge(usage_summary, inventory_summary, on=["Technician", "Part #", "Part Description"], how="outer")

# Fill missing values
merged["Quantity Used"] = merged["Quantity Used"].fillna(0)
merged["QoH"] = merged["QoH"].fillna(0)

# Calculate suggested order quantity
merged["Suggested Order Quantity"] = merged["Quantity Used"] - merged["QoH"]
merged["Suggested Order Quantity"] = merged["Suggested Order Quantity"].apply(lambda x: max(x, 0))

# Flag missing and used
merged["Missing and Used"] = (merged["Quantity Used"] > 0) & (merged["QoH"] == 0)

# Save to Excel
merged.to_excel("Cleaned_Technician_Inventory_Suggestions.xlsx", index=False)
print("Report generated successfully.")
