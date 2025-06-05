import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter

# File paths
csv_file = 'Adaptive DUM KPI History_05-2025.csv'
output_excel = 'Adaptive DUM KPI History_05-2025_dealergroups.xlsx'

# Load CSV
df = pd.read_csv(csv_file, dtype=str)

# Book3.xlsx format info
book3_columns = [
    'PARENT_ACCOUNT', 'LEVEL', 'EOM_DATE', 'ID', 'NAME', 'NEW_DUM', 'CHURN_DUM', 'FM_SUBSCRIPTION',
    'CCD_SUBSCRIPTION', 'TOLLS_SUBSCRIPTION', 'TOLLS_USAGE', 'TELEMATICS_SUBSCRIPTION', 'INSV_SUBSCRIPTION',
    'BOOKING_SUBSCRIPTION', 'ACTIVE_FM', 'ACTIVE_CCD', 'ACTIVE_TOLLS', 'ACTIVE_TELEMATICS', 'ACTIVE_SCHEDULER',
    'ACTIVE_INSV', 'ACTIVE_BOOKING', 'FM_CMRR', 'CCD_CMRR', 'TOLLS_CMRR', 'TELEMATICS_CMRR', 'INSV_CMRR',
    'BOOKING_CMRR', 'TOTAL_CMRR', 'COUNTRY', 'STATE', None, 'PRIMARY_MANUFACTURER', 'PARENT_ACCOUNT', None,
    'HIGH_TOLL_ZONE', 'EOM_VUM', 'UNIQUE_VUM'
]

import numpy as np
from datetime import datetime, timedelta

# Convert Excel serial numbers in EOM_DATE to date strings
def excel_serial_to_date(val):
    try:
        serial = int(val)
        if 30000 < serial < 50000:
            return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime('%Y-%m-%d')
        else:
            return val
    except:
        return val

df['EOM_DATE'] = df['EOM_DATE'].apply(excel_serial_to_date)

# Filter to only EOM_DATE in May 2025 (accepts '2025-05', '5/2025', or '2025-05-31')
filtered = df[
    df['EOM_DATE'].str.contains('2025-05', na=False) |
    df['EOM_DATE'].str.contains('5/2025', na=False) |
    df['EOM_DATE'].str.contains('2025-05-31', na=False)
].copy()

# Set PARENT_ACCOUNT to '0' where missing, empty, or 'nan'
if 'PARENT_ACCOUNT' in filtered.columns:
    filtered['PARENT_ACCOUNT'] = filtered['PARENT_ACCOUNT'].replace([None, '', np.nan, 'nan'], '0')
    filtered['PARENT_ACCOUNT'] = filtered['PARENT_ACCOUNT'].fillna('0')

# Reorder columns and add empty columns as needed
ordered_cols = []
for col in book3_columns:
    if col is None:
        ordered_cols.append(f'_EMPTY_{len(ordered_cols)}')
    else:
        ordered_cols.append(col)

# Create DataFrame with correct order and empty columns
out_df = pd.DataFrame()
for col in ordered_cols:
    if col.startswith('_EMPTY_'):
        out_df[col] = ''
    else:
        out_df[col] = filtered[col] if col in filtered.columns else ''

# Save to Excel with correct sheet name
sheet_name = 'April Contracted Data'
out_df.to_excel(output_excel, index=False, sheet_name=sheet_name)

# Apply autofilter (B1:AK2013)
wb = openpyxl.load_workbook(output_excel)
ws = wb[sheet_name]
ws.auto_filter.ref = 'B1:AK2013'
wb.save(output_excel)

print(f"Filtered dealer groups written to {output_excel} with Book3.xlsx format")
