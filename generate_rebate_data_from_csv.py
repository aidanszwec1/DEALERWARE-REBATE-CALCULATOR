import pandas as pd

# --- Config ---
INPUT_FILE = 'Adaptive DUM KPI History_05-2025.csv'
OUTPUT_FILE = 'Rebate_data.xlsx'
SERVICED_BRANDS = [
    "Audi", "Ford", "Lamborghini", "Jaguar", "Landrover", "Land Rover", "Lexus", "Lincoln", "Toyota", "Volkswagen", "Volvo",
    "Nissan", "Infiniti", "Acura", "Mercedes"
]
MONTH_FILTER = '2025-05'

# --- Load Data ---
df = pd.read_csv(INPUT_FILE, dtype=str)
df['EOM_DATE'] = df['EOM_DATE'].astype(str)
df = df[df['EOM_DATE'].str.contains(MONTH_FILTER)]

# Ensure numeric columns
df['EOM_VUM'] = pd.to_numeric(df['EOM_VUM'], errors='coerce').fillna(0)
df['TOTAL_CMRR'] = pd.to_numeric(df['TOTAL_CMRR'], errors='coerce').fillna(0)

# --- Group and Calculate ---
groups = []
for group, group_df in df.groupby('PARENT_ACCOUNT'):
    # N&I: blank (if needed later)
    n_i = ""
    # TAM: sum of EOM_VUM for all stores
    tam = group_df['EOM_VUM'].sum()
    # SAM: sum of EOM_VUM for serviced brands
    sam_mask = group_df['PRIMARY_MANUFACTURER'].isin(SERVICED_BRANDS)
    sam = group_df.loc[sam_mask, 'EOM_VUM'].sum()
    # DUM: count of all stores
    dum = group_df.shape[0]
    # NUM_SAM: count of stores in SAM
    num_sam = group_df.loc[sam_mask].shape[0]
    # MRR: sum of TOTAL_CMRR for all stores
    mrr = group_df['TOTAL_CMRR'].sum()
    # ARR: MRR * 12
    arr = mrr * 12
    # SAM Penetration: DUM / SAM
    sam_pen = sam / tam if tam else 0

    # Group Name (use first available)
    group_name = group_df['PARENT_ACCOUNT'].iloc[0]

    groups.append({
        "Group Name | Website": group_name,
        "TAM": int(tam),
        "SAM": int(sam),
        "DUM": int(dum),
        "NUM_SAM": int(num_sam),
        "SAM PEN": round(sam_pen, 6),
        "MRR": int(mrr),
        "Net New ARR": int(arr)
    })

# --- Output ---
summary_df = pd.DataFrame(groups)
summary_df.to_excel(OUTPUT_FILE, index=False)
print(f"Dealer group summary written to {OUTPUT_FILE}")
