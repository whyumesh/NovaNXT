import pandas as pd
import os

# === CONFIGURATION ===
input_file = "NovaNXT Rx-Oct'25.csv"  # Input CSV file
output_folder = "ZBM_Files"           # Folder where output files will be saved

# === READ CSV FILE ===
df = pd.read_csv(input_file)

# Clean column names (remove extra spaces if any)
df.columns = df.columns.str.strip()

# Ensure required columns exist
required_cols = [
    "ZBM Code", "ZBM Name", "ABM Code", "ABM Name",
    "Territory Code", "User: Full Name", "Account: Customer Code"
]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    raise ValueError(f"Missing columns in CSV: {missing}")

# Create output folder if it doesn’t exist
os.makedirs(output_folder, exist_ok=True)

# Group by ZBM Code and Name
grouped = df.groupby(["ZBM Code", "ZBM Name"], dropna=False)

for (zbm_code, zbm_name), group in grouped:
    # Reorder columns as per your desired format
    output_df = group[[
        "ZBM Code", "ZBM Name", "ABM Code", "ABM Name",
        "Territory Code", "User: Full Name", "Account: Customer Code"
    ]]

    # Create clean filename
    safe_zbm_name = str(zbm_name).replace("/", "_").replace("\\", "_").replace(":", "_")
    output_path = os.path.join(output_folder, f"ZBM_{zbm_code}_{safe_zbm_name}.xlsx")

    # Save each ZBM file separately
    output_df.to_excel(output_path, index=False)

print(f"✅ Files created successfully in '{output_folder}' folder!")
