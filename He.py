# === STEP 2: Find top brand and corresponding Rx/Month for each Account ===
brand_rx_pairs = [(f"Brand{i}: Brand Code", f"Rx/Month{i}") for i in range(1, 11)]

# Ensure numeric Rx columns are treated as numbers
for _, rx_col in brand_rx_pairs:
    if rx_col in df.columns:
        df[rx_col] = pd.to_numeric(df[rx_col], errors="coerce")

# Prepare a DataFrame with Account Code and its brand columns
brand_df = df[[mapped_cols["Account: Customer Code"]] + sum(brand_rx_pairs, ())].copy()

def get_top_brand_and_rx(row):
    top_brand, top_rx = None, float("-inf")
    for brand_col, rx_col in brand_rx_pairs:
        if brand_col in row and rx_col in row:
            rx_val = pd.to_numeric(row[rx_col], errors="coerce")
            if pd.notna(rx_val) and rx_val > top_rx:
                top_rx = rx_val
                top_brand = row[brand_col]
    return pd.Series([top_brand, top_rx], index=["Top Brand Code", "Top Rx/Month"])

# Apply function per row
brand_df[["Top Brand Code", "Top Rx/Month"]] = brand_df.apply(get_top_brand_and_rx, axis=1)

# Merge results with df_clean
df_clean = df_clean.merge(
    brand_df[[mapped_cols["Account: Customer Code"], "Top Brand Code", "Top Rx/Month"]],
    on=mapped_cols["Account: Customer Code"],
    how="left"
)
