import pandas as pd
import os

# ---------- CONFIG ----------
INPUT_FILE = r"NovaNXT Rx-Oct'25.csv"
OUTPUT_FOLDER = "ZBM_Files"
MASTER_SUMMARY_FILE = "All_ZBMs_Brand_Summary.xlsx"

# Create output directory
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------- LOAD DATA ROBUSTLY ----------
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

# ---------- CLEAN COLUMN NAMES ----------
df.columns = [c.strip() for c in df.columns]

# Rename key hierarchy columns
rename_map = {
    "ZBM Code": "ZBM Code",
    "ZBM Name": "ZBM Name",
    "ABM Code": "ABM Code",
    "ABM Name": "ABM Name",
    "Territory Code": "TBM Code",
    "User: Full Name": "TBM Name",
    "Account: Customer Code": "Dr Code"
}
df.rename(columns=rename_map, inplace=True)

# ---------- IDENTIFY BRAND COLUMNS ----------
brand_cols = [c for c in df.columns if "Brand" in c and "Code" in c]
rx_cols = [c for c in df.columns if "Rx/Month" in c]

# Check if pairs match
brand_rx_pairs = list(zip(brand_cols, rx_cols))

# ---------- EXPAND BRANDS INTO LONG FORMAT ----------
records = []
for _, row in df.iterrows():
    for brand_col, rx_col in brand_rx_pairs:
        brand = row.get(brand_col)
        rx = row.get(rx_col)
        if pd.notna(brand):
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
brand_df["Rx/Month"] = pd.to_numeric(brand_df["Rx/Month"], errors="coerce").fillna(0)

# ---------- CREATE PER-ZBM FILES (Option 1) ----------
for zbm, group in brand_df.groupby(["ZBM Code", "ZBM Name"]):
    zbm_code, zbm_name = zbm
    file_path = os.path.join(OUTPUT_FOLDER, f"ZBM_{zbm_code}_{zbm_name}.xlsx")

    # Summary per level
    summary = group.groupby(["ZBM Code", "TBM Code", "ABM Code"], as_index=False).agg(
        Total_Rx=("Rx/Month", "sum"),
        Unique_Brands=("Brand", "nunique"),
        Doctors=("Dr Code", "nunique")
    )

    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        group.to_excel(writer, index=False, sheet_name="Brand Summary")
        summary.to_excel(writer, index=False, sheet_name="Summary")

    print(f"✅ Created {file_path}")

# ---------- CREATE MASTER FILE (Option 2) ----------
with pd.ExcelWriter(MASTER_SUMMARY_FILE, engine="xlsxwriter") as writer:
    brand_df.to_excel(writer, index=False, sheet_name="All Data")
    overall_summary = brand_df.groupby(
        ["ZBM Code", "ZBM Name", "TBM Code", "ABM Code", "Brand"], as_index=False
    ).agg(Total_Rx=("Rx/Month", "sum"))
    overall_summary.to_excel(writer, index=False, sheet_name="Summary")

print(f"✅ Master summary saved to {MASTER_SUMMARY_FILE}")
