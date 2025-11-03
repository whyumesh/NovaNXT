import pandas as pd
import os

# ---------- CONFIG ----------
INPUT_FILE = r"NovaNXT Rx-Oct'25.csv"
OUTPUT_FOLDER = "ZBM_Files_Brandwise"
MASTER_FILE = "All_ZBMs_Brandwise_Summary.xlsx"

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

# ---------- RENAME CORE COLUMNS ----------
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

# ---------- DETECT BRAND–RX PAIRS ----------
brand_cols = [c for c in df.columns if "Brand" in c and "Code" in c]
rx_cols = [c for c in df.columns if "Rx/Month" in c]
brand_rx_pairs = list(zip(brand_cols, rx_cols))

if not brand_rx_pairs:
    raise ValueError("❌ No Brand–Rx/Month pairs found. Please check CSV headers.")

# ---------- CLEAN TYPES ----------
df["Dr Code"] = df["Dr Code"].astype(str).str.strip()

# ---------- GROUP BY HIERARCHY ----------
hierarchy_cols = ["ZBM Code", "ZBM Name", "TBM Code", "TBM Name", "ABM Code", "ABM Name"]
grouped = df.groupby(hierarchy_cols)

# ---------- AGGREGATE BRANDWISE ----------
rows = []

for keys, group in grouped:
    record = dict(zip(hierarchy_cols, keys))
    
    # for each brand/rx column pair
    for i, (brand_col, rx_col) in enumerate(brand_rx_pairs, start=1):
        valid = group[pd.notna(group[brand_col])]
        if valid.empty:
            record[f"Brand{i}: Brand Code"] = None
            record[f"Rx/Month{i}"] = 0
            continue
        
        # get unique brands and their Rx totals
        brands = valid[brand_col].dropna().unique()
        rx_sum = valid.groupby(brand_col)[rx_col].sum(min_count=1)
        dr_counts = valid.groupby(brand_col)["Dr Code"].nunique()

        # if multiple brands exist in same column, pick top one by Rx sum
        top_brand = rx_sum.sort_values(ascending=False).index[0]
        record[f"Brand{i}: Brand Code"] = top_brand
        record[f"Rx/Month{i}"] = rx_sum[top_brand]
        record[f"Unique Drs {i}"] = dr_counts[top_brand]
    
    rows.append(record)

summary_df = pd.DataFrame(rows)

# ---------- SAVE PER ZBM (Option 1) ----------
for (zbm_code, zbm_name), zbm_df in summary_df.groupby(["ZBM Code", "ZBM Name"]):
    file_path = os.path.join(OUTPUT_FOLDER, f"ZBM_{zbm_code}_{zbm_name}_Brandwise.xlsx")
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        zbm_df.to_excel(writer, index=False, sheet_name="Brand Summary")
    print(f"✅ Created {file_path}")

# ---------- SAVE MASTER SUMMARY (Option 2) ----------
summary_df.to_excel(MASTER_FILE, index=False)
print(f"✅ Master summary saved to {MASTER_FILE}")
