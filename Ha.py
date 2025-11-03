import pandas as pd
import os

# ---------- CONFIG ----------
INPUT_FILE = r"NovaNXT Rx-Oct'25.csv"
OUTPUT_FOLDER = "ZBM_Files_WideSummary"
MASTER_FILE = "All_ZBMs_Wide_Summary.xlsx"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------- LOAD DATA ROBUSTLY ----------
encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "iso-8859-1", "cp850"]
for enc in encodings:
    try:
        df = pd.read_csv(INPUT_FILE, encoding=enc)
        print(f"✅ Loaded CSV successfully with encoding: {enc}")
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

# ---------- DETECT BRAND-RX COLUMN PAIRS ----------
brand_cols = [c for c in df.columns if "Brand" in c and "Code" in c]
rx_cols = [c for c in df.columns if "Rx/Month" in c]
brand_rx_pairs = list(zip(brand_cols, rx_cols))

if not brand_rx_pairs:
    raise ValueError("❌ No Brand–Rx/Month pairs found. Please check CSV headers.")

# ---------- CLEAN TYPES ----------
df["Dr Code"] = df["Dr Code"].astype(str).str.strip()

# ---------- HIERARCHY GROUP ----------
hierarchy_cols = ["ZBM Code", "ZBM Name", "TBM Code", "TBM Name", "ABM Code", "ABM Name"]

# ---------- CREATE SUMMARY ----------
summary_rows = []
for group_keys, group_df in df.groupby(hierarchy_cols):
    summary = dict(zip(hierarchy_cols, group_keys))
    
    for brand_col, rx_col in brand_rx_pairs:
        brand_name = brand_col.split(":")[0].strip()  # e.g., "Brand1"
        unique_doctors = group_df.loc[group_df[brand_col].notna(), "Dr Code"].nunique()
        total_rx = pd.to_numeric(group_df[rx_col], errors="coerce").sum()
        
        summary[brand_col] = unique_doctors
        summary[rx_col] = total_rx
    
    summary_rows.append(summary)

summary_df = pd.DataFrame(summary_rows)

# ---------- SAVE PER ZBM (Option 1) ----------
for (zbm_code, zbm_name), zbm_df in summary_df.groupby(["ZBM Code", "ZBM Name"]):
    file_path = os.path.join(OUTPUT_FOLDER, f"ZBM_{zbm_code}_{zbm_name}_WideSummary.xlsx")
    zbm_df.to_excel(file_path, index=False)
    print(f"✅ Created {file_path}")

# ---------- SAVE MASTER FILE (Option 2) ----------
summary_df.to_excel(MASTER_FILE, index=False)
print(f"✅ Master summary saved to {MASTER_FILE}")
