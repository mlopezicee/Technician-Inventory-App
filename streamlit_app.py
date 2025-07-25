import streamlit as st
import pandas as pd

st.title("Technician Inventory Suggestion Tool")

# Upload files
part_file = st.file_uploader("Upload Part History Report", type=["xlsx"])
inventory_file = st.file_uploader("Upload Inventory Report", type=["xlsx"])

# Define expected columns
expected_columns = {
    "Technician": ["Technician", "Tech Name", "Tech"],
    "Part #": ["Part #", "Part Number", "Part Num"],
    "Part Description": ["Part Description", "Description"],
    "Quantity Used": ["Quantity Used", "Qty", "Qty Used"],
    "QoH": ["QoH", "Quantity on Hand", "On Hand"]
}

def match_columns(df, expected_map):
    matched = {}
    for key, options in expected_map.items():
        for col in df.columns:
            if any(opt.lower() in str(col).lower() for opt in options):
                matched[key] = col
                break
    return matched

if part_file and inventory_file:
    try:
        part_df = pd.read_excel(part_file, engine="openpyxl", header=2)
        inventory_df = pd.read_excel(inventory_file, engine="openpyxl", header=1)

        part_cols = match_columns(part_df, expected_columns)
        inv_cols = match_columns(inventory_df, expected_columns)

        required = ["Technician", "Part #", "Part Description", "Quantity Used", "QoH"]
        if not all(col in part_cols for col in ["Technician", "Part #", "Quantity Used"]) or            not all(col in inv_cols for col in ["Technician", "Part #", "Part Description", "QoH"]):
            st.error("Missing required columns in one or both files.")
        else:
            part_df["Part #"] = part_df[part_cols["Part #"]].astype(str).str.strip()
            inventory_df["Part #"] = inventory_df[inv_cols["Part #"]].astype(str).str.strip()

            usage_summary = part_df.groupby(
                [part_cols["Technician"], "Part #"]
            )[part_cols["Quantity Used"]].sum().reset_index()
            usage_summary.columns = ["Technician", "Part #", "Quantity Used"]

            inventory_df["Technician"] = inventory_df[inv_cols["Technician"]]
            inventory_df["Part Description"] = inventory_df[inv_cols["Part Description"]]
            inventory_df["QoH"] = inventory_df[inv_cols["QoH"]]

            merged = pd.merge(
                usage_summary,
                inventory_df[["Technician", "Part #", "Part Description", "QoH"]],
                on=["Technician", "Part #"],
                how="left"
            )

            merged["QoH"] = merged["QoH"].fillna(0)
            merged["Suggested Order Quantity"] = merged["Quantity Used"] - merged["QoH"]
            merged["Suggested Order Quantity"] = merged["Suggested Order Quantity"].apply(lambda x: max(x, 0))
            merged["Missing and Used"] = (merged["QoH"] == 0) & (merged["Quantity Used"] > 0)

            st.success("Suggested parts list generated.")
            st.dataframe(merged)

            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                merged.to_excel(writer, index=False, sheet_name="Sheet1")
            st.download_button(
                label="Download Suggested Parts Report",
                data=output.getvalue(),
                file_name="Cleaned_Technician_Inventory_Suggestions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error processing files: {e}")