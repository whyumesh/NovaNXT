import pandas as pd
import os, re, zipfile

# === CONFIGURATION ===
input_file = "NovaNXT Rx-Oct'25.csv"
output_folder = "ZBM_Files"
zip_file = "ZBM_Files.zip"
master_summary_file = "Master_Brand_Summary.xlsx"

# --- 1️⃣  Robust CSV Reader ---
def read_csv_robust(filepath):
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1", "iso-8859-1", "cp850"]
    for enc in encodings:
        try:
            df = pd.read_csv(filepath, encoding=enc, dtype=str, low_memory=False)
            print(f"✅ Read success with {enc}, rows={len(df)}")
            return df
        except Exception as e:
            print(f"⚠️ {enc} failed: {e}")
    df = pd.read_csv(filepath, encoding="latin-1", dtype=str, low_memory=False, errors="replace")
    print("✅ Fallback read with latin-1 (errors replaced)")
    return df

df = read_csv_robust(input_file)
df.columns = df.columns.str.strip()

# --- 2️⃣  Column mapping for hierarchy ---
column_map = {
    "ZBM Code": ["ZBM Code","zbm_code"],
    "ZBM Name": ["ZBM Name","zbm_name"],
    "ABM Code": ["ABM Code","abm_code"],
    "ABM Name": ["ABM Name","abm_name"],
    "Territory Code": ["Territory Code","tbm_code"],
    "User: Full Name": ["User: Full Name","tbm_name"],
    "Account: Customer Code": ["Account: Customer Code","Dr Code","doctor_code"]
}
def find_col(opts):
    for o in opts:
        if o in df.columns: return o
    return None
mapped = {k: find_col(v) for k,v in column_map.items()}
if None in mapped.values():
    missing = [k for k,v in mapped.items() if v is None]
    raise ValueError(f"Missing required columns: {missing}")

# --- 3️⃣  Brand column detection ---
brand_cols = []
for i in range(1,11):
    code_col = f"Brand{i}: Brand Code"
    rx_col   = f"Rx/Month{i}"
    if code_col in df.columns and rx_col in df.columns:
        brand_cols.append((code_col, rx_col))
if not brand_cols:
    raise ValueError("No brand columns found (expected Brand1..Brand10).")

# --- 4️⃣  Clean hierarchy data ---
hier_cols = list(mapped.keys())
base_df = pd.DataFrame({k: df[mapped[k]].astype(str).fillna("").str.strip() for k in hier_cols})

# --- 5️⃣  Build brand summary ---
brand_records = []
for idx, row in base_df.iterrows():
    zbm_c, zbm_n, abm_c, abm_n, tbm_c, tbm_n, dr = \
        row["ZBM Code"], row["ZBM Name"], row["ABM Code"], row["ABM Name"], \
        row["Territory Code"], row["User: Full Name"], row["Account: Customer Code"]
    for (bcol, rxcol) in brand_cols:
        brand = str(df.at[idx, bcol]).strip()
        rx = df.at[idx, rxcol]
        try:
            rx_val = float(rx) if str(rx).strip() not in ["", "nan"] else 0.0
        except:
            rx_val = 0.0
        if brand != "" or rx_val != 0:
            brand_records.append({
                "ZBM Code": zbm_c,
                "ZBM Name": zbm_n,
                "ABM Code": abm_c,
                "ABM Name": abm_n,
                "Territory Code": tbm_c,
                "User: Full Name": tbm_n,
                "Account: Customer Code": dr,
                "Brand Code": brand,
                "Rx/Month": rx_val
            })

brand_df = pd.DataFrame(brand_records)

# --- 6️⃣  Create per-ZBM files with brand summary (Option 1) ---
os.makedirs(output_folder, exist_ok=True)
created = []

grouped = base_df.groupby(["ZBM Code","ZBM Name"], dropna=False)
for (zbm_code, zbm_name), grp in grouped:
    zbm_brands = brand_df.query("`ZBM Code` == @zbm_code and `ZBM Name` == @zbm_name")

    # --- Summary counts ---
    summary = pd.DataFrame({
        "Metric": [
            "Total Rows",
            "Unique TBMs",
            "Unique ABMs",
            "Unique Doctors",
            "Total Brand Entries",
            "Total Rx (Sum)"
        ],
        "Value": [
            len(grp),
            grp["Territory Code"].nunique(),
            grp["ABM Code"].nunique(),
            grp["Account: Customer Code"].nunique(),
            len(zbm_brands),
            zbm_brands["Rx/Month"].sum()
        ]
    })

    # --- Brand summary (per ABM) ---
    brand_summary = (
        zbm_brands.groupby(["ABM Code","ABM Name","Territory Code","User: Full Name"], dropna=False)
        .agg(Brands_Handled=("Brand Code", lambda x: x.notna().sum()),
             Total_Rx=("Rx/Month","sum"))
        .reset_index()
    )

    # --- Save to Excel ---
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", f"ZBM_{zbm_code}_{zbm_name}")[:150]
    path = os.path.join(output_folder, f"{safe_name}.xlsx")

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        grp.to_excel(w, sheet_name="Data", index=False)
        summary.to_excel(w, sheet_name="Summary", index=False)
        brand_summary.to_excel(w, sheet_name="Brand Summary", index=False)

    created.append(path)
    print(f"✅ Created: {path}")

# --- 7️⃣  Master Brand Summary (Option 2) ---
master_summary = (
    brand_df.groupby(["ZBM Code","ZBM Name","ABM Code","ABM Name","Territory Code","User: Full Name"], dropna=False)
    .agg(Brands_Handled=("Brand Code", lambda x: x.notna().sum()),
         Total_Rx=("Rx/Month","sum"))
    .reset_index()
)
master_summary.to_excel(master_summary_file, index=False)
print(f"📘 Master Brand Summary saved: {master_summary_file}")

# --- 8️⃣  Zip all ZBM files ---
with zipfile.ZipFile(zip_file, "w", compression=zipfile.ZIP_DEFLATED) as zf:
    for f in created:
        zf.write(f, arcname=os.path.basename(f))
print(f"📦 All ZBM files zipped → {zip_file}")
