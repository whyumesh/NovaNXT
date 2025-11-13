import pandas as pd

# === CONFIGURATION ===
input_file = "Rx_prescription_3Months.csv"  # Single file containing all months
output_file = "ZBM_Combined_3Months_Report.xlsx"

# === FUNCTION TO READ CSV ROBUSTLY ===
def read_csv_robust(filepath):
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "iso-8859-1", "cp850"]
    for enc in encodings:
        try:
            df = pd.read_csv(filepath, encoding=enc, dtype=str, low_memory=False)
            print(f"✅ Successfully read {filepath} with encoding: {enc} — Rows: {len(df)}")
            return df
        except Exception as e:
            print(f"⚠️ Failed with encoding {enc}: {e}")
    # Fallback read
    df = pd.read_csv(filepath, encoding="latin-1", dtype=str, low_memory=False)
    print(f"✅ Fallback read with latin-1 for {filepath}")
    return df

# === COLUMN MAPPING ===
column_map = {
    "Date": ["Date", "date"],
    "Division": ["Division"],
    "Territory Code": ["Territory Code"],
    "User: Full Name": ["User: Full Name"],
    "Account: Customer Code": ["Account: Customer Code"]
}

def process_monthly_data(df):
    df.columns = df.columns.str.strip()

    def find_column(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None

    mapped_cols = {}
    for final_name, options in column_map.items():
        col = find_column(options)
        if not col:
            raise ValueError(f"Missing required column: {final_name}")
        mapped_cols[final_name] = col

    df_clean = pd.DataFrame()
    for final_name, original in mapped_cols.items():
        df_clean[final_name] = df[original].astype(str).fillna("").str.strip()

    # Handle Brand/Rx columns dynamically
    brand_rx_pairs = [(f"Brand{i}: Brand Code", f"Rx/Month{i}") for i in range(1, 11)]
    existing_pairs = [(b, r) for b, r in brand_rx_pairs if b in df.columns and r in df.columns]

    for _, rx_col in existing_pairs:
        df[rx_col] = pd.to_numeric(df[rx_col], errors="coerce")

    for brand_col, rx_col in existing_pairs:
        df_clean[brand_col] = df[brand_col].astype(str).fillna("").str.strip()
        df_clean[rx_col] = df[rx_col]

    # Group by Account and pick max Rx per brand
    grouped_data = []
    for account_code, group in df_clean.groupby("Account: Customer Code"):
        base_row = group.iloc[0][list(mapped_cols.keys())].to_dict()
        for brand_col, rx_col in existing_pairs:
            max_idx = group[rx_col].idxmax()
            if pd.notna(max_idx) and max_idx in group.index:
                rx_value = group.loc[max_idx, rx_col]
                base_row[brand_col] = group.loc[max_idx, brand_col] if pd.notna(rx_value) else None
                base_row[rx_col] = rx_value if pd.notna(rx_value) else None
            else:
                base_row[brand_col] = None
                base_row[rx_col] = None
        grouped_data.append(base_row)

    df_full = pd.DataFrame(grouped_data)
    hierarchy_cols = list(mapped_cols.keys())
    brand_rx_ordered = [col for pair in existing_pairs for col in pair]
    final_column_order = hierarchy_cols + brand_rx_ordered
    df_full = df_full[final_column_order]
    df_full = df_full.sort_values(by=["Division", "Territory Code", "User: Full Name"], na_position='last').reset_index(drop=True)
    return df_full

# === READ SINGLE FILE ===
df_all = read_csv_robust(input_file)

# ✅ Convert Date column to datetime
date_col = [c for c in df_all.columns if c.lower() == "date"]
if not date_col:
    raise ValueError("❌ Date column is missing in the file!")
date_col = date_col[0]

df_all[date_col] = pd.to_datetime(df_all[date_col], format="%d-%m-%y", errors="coerce")
df_all["Month"] = df_all[date_col].dt.strftime("%b")  # Sep, Oct, Nov

# Split by Month
monthly_data = {}
for month in df_all["Month"].unique():
    df_month = df_all[df_all["Month"] == month]
    monthly_data[month] = process_monthly_data(df_month)

# Combine with spacers
max_rows = max(len(m) for m in monthly_data.values())
aligned_months = [df.reindex(range(max_rows)).reset_index(drop=True) for df in monthly_data.values()]
spacer_cols = [f"Spacer{i}" for i in range(1, 11)]
spacer_df = pd.DataFrame(columns=spacer_cols, index=range(max_rows))

combined_df = pd.concat([aligned_months[0], spacer_df, aligned_months[1], spacer_df, aligned_months[2]], axis=1)

# Write to Excel
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    combined_df.to_excel(writer, sheet_name="Combined Report", index=False)

print(f"\n🎉 Combined report generated successfully!")
print(f"📄 Output file: {output_file}")
print(f"📊 Months processed: {list(monthly_data.keys())}")
