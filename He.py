Here's the modified code that creates a single Excel file with separate sheets for each ZBM:
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

# === GROUP BY ZBM CODE + NAME ===
grouped = df_full.groupby(["ZBM Code", "ZBM Name"], dropna=False)

# === CREATE SINGLE EXCEL FILE WITH MULTIPLE SHEETS ===
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    
    # Create an overall summary sheet
    overall_summary = pd.DataFrame({
        "ZBM Code": [],
        "ZBM Name": [],
        "Total Rows": [],
        "Unique TBM": [],
        "Unique ABM": [],
        "Unique Doctors": []
    })
    
    summary_data = []
    sheet_count = 0
    
    for (zbm_code, zbm_name), group in grouped:
        group = group.drop_duplicates().reset_index(drop=True)
        
        # Create safe sheet name (Excel limits: 31 chars, no special chars)
        safe_sheet_name = re.sub(r'[\\/*?\[\]:"]', "_", f"{zbm_code}_{zbm_name}")[:31]
        
        # Ensure unique sheet name
        original_name = safe_sheet_name
        counter = 1
        while safe_sheet_name in writer.sheets:
            safe_sheet_name = f"{original_name[:28]}_{counter}"
            counter += 1
        
        # Write data to sheet
        group.to_excel(writer, sheet_name=safe_sheet_name, index=False)
        
        # Collect summary data
        summary_data.append({
            "ZBM Code": zbm_code,
            "ZBM Name": zbm_name,
            "Total Rows": len(group),
            "Unique TBM": group["Territory Code"].nunique(),
            "Unique ABM": group["ABM Code"].nunique(),
            "Unique Doctors": group["Account: Customer Code"].nunique()
        })
        
        sheet_count += 1
        print(f"✅ Added sheet: {safe_sheet_name}")
    
    # Write overall summary as first sheet
    overall_summary = pd.DataFrame(summary_data)
    overall_summary.to_excel(writer, sheet_name="Overall Summary", index=False)
    
    # Move summary sheet to first position
    workbook = writer.book
    summary_sheet = workbook["Overall Summary"]
    workbook.move_sheet(summary_sheet, offset=-sheet_count)

print(f"\n🎉 All done! Created {sheet_count} sheets in one file.")
print(f"📄 Output file: {output_file}")
Key changes:
Single file output: Creates one Excel file (ZBM_Combined_Report.xlsx) instead of multiple files
Multiple sheets: Each ZBM gets its own sheet within the file
Overall Summary: Added a summary sheet (placed first) showing metrics for all ZBMs
No ZIP file: Since everything is in one file, no need for zipping
Sheet name handling: Excel sheet names are limited to 31 characters and can't contain special characters - the code handles this automatically
The file will contain:
Overall Summary sheet: Lists all ZBMs with their row counts and unique counts
One sheet per ZBM: Contains the filtered data for that ZBM
