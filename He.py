Here's the updated code with columns arranged as Brand → Rx for each brand pair:

```python
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

# === FIND HIGHEST Rx/MONTH FOR EACH ACCOUNT ===
def get_highest_brand_rx(row):
    max_val = float("-inf")
    top_brand = None
    for brand_col, rx_col in existing_pairs:
        val = row.get(rx_col, None)
        if pd.notna(val) and val > max_val:
            max_val = val
            top_brand = row.get(brand_col, None)
    result = {brand_col: None for brand_col, _ in existing_pairs}
    result.update({rx_col: None for _, rx_col in existing_pairs})
    if pd.notna(top_brand):
        # find which brand column had that brand and assign only that
        for brand_col, rx_col in existing_pairs:
            if row.get(brand_col) == top_brand and row.get(rx_col) == max_val:
                result[brand_col] = top_brand
                result[rx_col] = max_val
                break
    return pd.Series(result)

# Apply per Account row
brand_rx_filtered = df.apply(get_highest_brand_rx, axis=1)

# Combine with main hierarchy dataframe
df_full = pd.concat([df_clean, brand_rx_filtered], axis=1)

# === REMOVE DUPLICATES ===
df_full = df_full.drop_duplicates().reset_index(drop=True)

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
    summary_df.to_excel(writer, sheet_name="Summary", index=False)

print(f"\n🎉 All done! All ZBM data combined in one sheet.")
print(f"📄 Output file: {output_file}")
print(f"📊 Total rows: {len(df_full)}")
print(f"📊 Total ZBMs: {df_full['ZBM Code'].nunique()}")
```

**Key changes:**

1. **Column ordering section added** after removing duplicates
2. **Hierarchy columns first**: ZBM Code, ZBM Name, ABM Code, ABM Name, Territory Code, User: Full Name, Account: Customer Code
3. **Brand-Rx pairs follow**: Brand1: Brand Code, Rx/Month1, Brand2: Brand Code, Rx/Month2, etc.
4. **DataFrame reordered** using the `final_column_order` list

The output will now have columns in this order:
- ZBM Code
- ZBM Name
- ABM Code
- ABM Name
- Territory Code
- User: Full Name
- Account: Customer Code
- Brand1: Brand Code
- Rx/Month1
- Brand2: Brand Code
- Rx/Month2
- ... (and so on for all brand-rx pairs)
