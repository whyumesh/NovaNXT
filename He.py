import pandas as pd
import os
import zipfile
import re

# === CONFIGURATION ===
input_file = "NovaNXT Rx-Oct'25.csv"  # Path to your CSV
output_folder = "ZBM_Files"            # Folder to store ZBM Excel files
zip_file = "ZBM_Files.zip"             # Final zip file name

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

# === CREATE OUTPUT FOLDER ===
os.makedirs(output_folder, exist_ok=True)

# === GROUP BY ZBM CODE + NAME ===
grouped = df_full.groupby(["ZBM Code", "ZBM Name"], dropna=False)

created_files = []

for (zbm_code, zbm_name), group in grouped:
    group = group.drop_duplicates().reset_index(drop=True)

    # === SUMMARY ===
    summary = pd.DataFrame({
        "Metric": [
            "Total Rows in File",
            "Unique TBM (Territory Code)",
            "Unique ABM Code",
            "Unique Doctor (Account: Customer Code)"
        ],
        "Value": [
            len(group),
            group["Territory Code"].nunique(),
            group["ABM Code"].nunique(),
            group["Account: Customer Code"].nunique()
        ]
    })

    # === SAFE FILE NAME ===
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", f"ZBM_{zbm_code}_{zbm_name}")[:150]
    output_path = os.path.join(output_folder, f"{safe_name}.xlsx")

    # === WRITE TO EXCEL ===
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        group.to_excel(writer, sheet_name="Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    created_files.append(output_path)
    print(f"✅ Created: {output_path}")

# === ZIP ALL FILES ===
with zipfile.ZipFile(zip_file, "w", compression=zipfile.ZIP_DEFLATED) as zf:
    for f in created_files:
        zf.write(f, arcname=os.path.basename(f))

print(f"\n🎉 All done! {len(created_files)} files created.")
print(f"📦 Zipped file: {zip_file}")
