import pandas as pd
import os

# ---------- CONFIG ----------
INPUT_FILE = r"NovaNXT Rx-Oct'25.csv"
OUTPUT_FOLDER = "ZBM_Files"
MASTER_SUMMARY_FILE = "All_ZBMs_Brand_Summary.xlsx"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------- LOAD CSV ROBUSTLY ----------
encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "iso-8859-1", "cp850"]
for enc in encodings:
    try:
        df = pd.read_csv(INPUT_FILE, encoding=enc)
        print(f"✅ Loaded CSV successfully using encoding: {enc}")
        break
    except Exception as e:
        print(f"❌ Failed with {enc}: {e}")
else:
    raise ValueError("Failed to read CSV with all tested encodings")

df.columns = [c.strip() for c in df.columns]

# ---------- RENAME MAIN COLUMNS ----------
rename_map = {
    "ZBM Code": "ZBM Code",
    "ZBM Name": "ZBM Name",
    "ABM Code": "ABM Code",
    "ABM Name": "ABM Name",
    "Territory Code": "TBM Code",
    "User: Full Name": "TBM Name",
    "Account: Customer Code": "Dr Code",
}
df.rename(columns=rename_map, inplace=True)

# ---------- DETECT BRAND-RX COLUMN PAIRS ----------
brand_cols = [c for c in df.columns if "Brand" in c and "Code" in c]
rx_cols = [c for c in df.columns if "Rx/Month" in c]
brand_rx_pairs = list(zip(brand_cols, rx_cols))

if not brand_rx_pairs:
    raise ValueError("❌ Could not detect any Brand–Rx/Month pairs. Please check column names.")

# ---------- EXPAND BRAND DATA INTO LONG FORMAT ----------
records = []
for _, row in df.iterrows():
    for brand_col, rx_col in brand_rx_pairs:
        brand = row.get(brand_col)
        rx = row.get(rx_col)
        if pd.notna(brand):  # only include valid brands
            records.append({
                "ZBM Code": row.get("ZBM Code"),
                "ZBM Name": row.get("ZBM Name"),
                "TBM Code": row.get("TBM Code"),
                "TBM Name": row.get("TBM Name"),
                "ABM Code": row.get("ABM Code"),
                "ABM Name": row.get("ABM Name"),
                "Dr Code": row.get("Dr Code"),
                "Brand": brand,
                "Rx/Month": rx
            })

brand_df = pd.DataFrame(records)

# ---------- ENSURE TYPES ----------
brand_df["Rx/Month"] = pd.to_numeric(brand_df["Rx/Month"], errors="coerce").fillna(0)
brand_df["Dr Code"] = brand_df["Dr Code"].astype(str).str.strip()

# ---------- PER-ZBM FILES (Option 1) ----------
for (zbm_code, zbm_name), group in brand_df.groupby(["ZBM Code", "ZBM Name"]):
    file_path = os.path.join(OUTPUT_FOLDER, f"ZBM_{zbm_code}_{zbm_name}.xlsx")

    # Create unique doctor count summary
    summary = (
        group.groupby(["ZBM Code", "TBM Code", "ABM Code", "Brand"], as_index=False)
        .agg(Unique_Doctors=("Dr Code", pd.Series.nunique))
        .sort_values(["TBM Code", "ABM Code", "Brand"])
    )

    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        group.to_excel(writer, index=False, sheet_name="Brand Data")
        summary.to_excel(writer, index=False, sheet_name="Unique Doctor Summary")

    print(f"✅ Created {file_path}")

# ---------- MASTER SUMMARY (Option 2) ----------
master_summary = (
    brand_df.groupby(["ZBM Code", "ZBM Name", "TBM Code", "ABM Code", "Brand"], as_index=False)
    .agg(Unique_Doctors=("Dr Code", pd.Series.nunique))
    .sort_values(["ZBM Code", "TBM Code", "ABM Code", "Brand"])
)

with pd.ExcelWriter(MASTER_SUMMARY_FILE, engine="xlsxwriter") as writer:
    brand_df.to_excel(writer, index=False, sheet_name="All Brand Data")
    master_summary.to_excel(writer, index=False, sheet_name="Unique Doctor Summary")

print(f"✅ Master summary saved to {MASTER_SUMMARY_FILE}")
