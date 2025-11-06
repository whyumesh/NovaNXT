import pandas as pd
import os
import re

# === CONFIGURATION ===
input_file = "NovaNXT Rx-Oct'25.csv"  # Path to your CSV
output_file = "ZBM_Combined_Report.xlsx"  # Single output file

# === FUNCTION TO READ CSV ROBUSTLY ===
def read_csv_robust(filepath):
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "iso-8859-1", "cp850"]
    for enc in encodings:
        try:
            df = pd.read_csv(filepath, encoding=enc, dtype=str, low_memory=False)
            print(f"✅ Successfully read with encoding: {enc} — Rows: {len(df)}")
            return df
        except Exception as e:
            print(f"⚠️ Failed with encoding {enc}: {e}")
    # Fallback read
    df = pd.read_csv(filepath, encoding="latin-1", dtype=str, low_memory=False, errors="replace")
    print("✅ Fallback read with latin-1 (errors replaced)")
    return df

# === READ INPUT FILE ===
df = read_csv_robust(input_file)
df.columns = df.columns.str.strip()  # Clean header spaces

# === COLUMN MAPPING ===
column_map = {
    "ZBM Code": ["ZBM Code", "zbm_code", "ZBMCode"],
    "ZBM Name": ["ZBM Name", "zbm_name", "ZBMName"],
    "ABM Code": ["ABM Code", "abm_code", "ABMCode"],
    "ABM Name": ["ABM Name", "abm_name", "ABMName"],
    "Territory Code": ["Territory Code", "territory_code", "TBM Code", "tbm_code"],
    "User: Full Name": ["User: Full Name", "user_full_name", "TBM Name", "tbm_name"],
    "Account: Customer Code": ["Account: Customer Code", "Dr Code", "doctor_code"]
}

def find_column(possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

mapped_cols = {}
missing = []
for final_name, options in column_map.items():
    col = find_column(options)
    if col:
        mapped_cols[final_name] = col
    else:
        missing.append(final_name)

if missing:
    raise ValueError(f"Missing required columns: {missing}")

# === CORE HIERARCHY DATA ===
df_clean = pd.DataFrame()
for final_name, original in mapped_cols.items():
    df_clean[final_name] = df[original].astype(str).fillna("").str.strip()

# === IDENTIFY BRAND & RX COLUMNS ===
brand_rx_pairs = [(f"Brand{i}: Brand Code", f"Rx/Month{i}") for i in range(1, 11)]
existing_pairs = [(b, r) for b, r in brand_rx_pairs if b in df.columns and r in df.columns]

# Convert Rx/Month columns to numeric
for _, rx_col in existing_pairs:
    df[rx_col] = pd.to_numeric(df[rx_col], errors="coerce")

# === PREPARE BRAND-RX DATA ===
# Add brand-rx columns to the clean dataframe
for brand_col, rx_col in existing_pairs:
    df_clean[brand_col] = df[brand_col].astype(str).fillna("").str.strip()
    df_clean[rx_col] = df[rx_col]

# === KEEP ONLY HIGHEST Rx FOR EACH ACCOUNT CUSTOMER CODE PER BRAND COLUMN ===
# Group by Account Customer Code and keep max Rx for each Brand/Rx pair
grouped_data = []

for account_code, group in df_clean.groupby("Account: Customer Code"):
    # Start with the first row's hierarchy data
    base_row = group.iloc[0][["ZBM Code", "ZBM Name", "ABM Code", "ABM Name", 
                               "Territory Code", "User: Full Name", "Account: Customer Code"]].to_dict()
    
    # For each Brand-Rx pair, find the maximum Rx value
    for brand_col, rx_col in existing_pairs:
        # Find row with max Rx for this brand column
        max_idx = group[rx_col].idxmax()
        if pd.notna(group.loc[max_idx, rx_col]):
            base_row[brand_col] = group.loc[max_idx, brand_col]
            base_row[rx_col] = group.loc[max_idx, rx_col]
        else:
            base_row[brand_col] = None
            base_row[rx_col] = None
    
    grouped_data.append(base_row)

# Create final dataframe
df_full = pd.DataFrame(grouped_data)

# === REORDER COLUMNS: Hierarchy first, then Brand-Rx pairs ===
hierarchy_cols = ["ZBM Code", "ZBM Name", "ABM Code", "ABM Name", 
                  "Territory Code", "User: Full Name", "Account: Customer Code"]

# Create ordered list: Brand1, Rx1, Brand2, Rx2, etc.
brand_rx_ordered = []
for brand_col, rx_col in existing_pairs:
    brand_rx_ordered.append(brand_col)
    brand_rx_ordered.append(rx_col)

# Final column order
final_column_order = hierarchy_cols + brand_rx_ordered

# Reorder dataframe
df_full = df_full[final_column_order]

# === SORT BY ZBM ===
# Sort by ZBM Code and ZBM Name to group all data for each ZBM together
df_full = df_full.sort_values(by=["ZBM Code", "ZBM Name", "ABM Code", "Territory Code"], 
                                na_position='last').reset_index(drop=True)

# === CREATE SINGLE EXCEL FILE ===
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    # Write all data in a single sheet
    df_full.to_excel(writer, sheet_name="All ZBM Data", index=False)
    
    # Create summary sheet
    grouped = df_full.groupby(["ZBM Code", "ZBM Name"], dropna=False)
    
    summary_data = []
    for (zbm_code, zbm_name), group in grouped:
        summary_data.append({
            "ZBM Code": zbm_code,
            "ZBM Name": zbm_name,
            "Total Rows": len(group),
            "Unique TBM": group["Territory Code"].nunique(),
            "Unique ABM": group["ABM Code"].nunique(),
            "Unique Doctors": group["Account: Customer Code"].nunique()
        })
    
    summary_df = pd.DataFrame(summary_data)
    # Sort summary by ZBM Code
    summary_df = summary_df.sort_values(by="ZBM Code", na_position='last').reset_index(drop=True)
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

print(f"\n🎉 All done! Data sorted and grouped by ZBM with unique accounts.")
print(f"📄 Output file: {output_file}")
print(f"📊 Total unique accounts: {len(df_full)}")
print(f"📊 Total ZBMs: {df_full['ZBM Code'].nunique()}")
print(f"📋 Each account shows highest Rx value for each brand column")
