import os
import pandas as pd
from glob import glob

from pdf_report_optimizer import generate_optimized_elevation_report

SALES_FOLDER = 'sales_data'
OUTPUT = 'sample_generated_report.pdf'

# Load sales files
files = sorted([os.path.join(SALES_FOLDER, f) for f in os.listdir(SALES_FOLDER) if f.lower().endswith('.csv') or f.lower().endswith('.xlsx')])
if not files:
    print('No sales files found in', SALES_FOLDER)
    raise SystemExit(1)

all_dfs = []
for f in files:
    try:
        df = pd.read_csv(f)
    except Exception:
        try:
            df = pd.read_excel(f)
        except Exception as e:
            print('Could not read', f, e)
            continue
    # extract sale no from filename
    import re
    m = re.search(r'Sale_(\d+)', os.path.basename(f))
    if m:
        sale_no = int(m.group(1))
        df['Sale_No'] = sale_no
    all_dfs.append(df)

if not all_dfs:
    print('No readable sales data')
    raise SystemExit(1)

data = pd.concat(all_dfs, ignore_index=True)

# pick latest sale
if 'Sale_No' not in data.columns:
    print('Sale_No column missing in data')
    raise SystemExit(1)

latest_sale = int(data['Sale_No'].max())
latest_df = data[data['Sale_No'] == latest_sale]

print('Loaded sales files:', files)
print('Latest sale:', latest_sale, 'rows in latest_df:', len(latest_df))

# Basic preprocessing to match dashboard expectations
if 'Status' in data.columns and 'Status_Clean' not in data.columns:
    data['Status_Clean'] = data['Status'].astype(str).str.strip().str.lower()

data['Total Weight'] = pd.to_numeric(data['Total Weight'], errors='coerce').fillna(0)
if 'Price' in data.columns:
    data['Price'] = pd.to_numeric(data['Price'], errors='coerce').fillna(0)

# Recompute latest_df after preprocessing
latest_df = data[data['Sale_No'] == latest_sale]

# Ensure categorical columns are strings and handle missing
for col in ['Sub Elevation', 'Grade', 'Broker', 'Buyer']:
    if col in data.columns:
        data[col] = data[col].astype(str).fillna('Unknown').replace('nan', 'Unknown')

# Recompute latest_df one more time after coercion
latest_df = data[data['Sale_No'] == latest_sale]

include_reports = {
    'report1': True,
    'report2': True,
    'report3': True,
    'report4': True,
    'report5': True,
}

print('Generating PDF...')
pdf_bytes = generate_optimized_elevation_report(data, latest_df, include_reports=include_reports)

with open(OUTPUT, 'wb') as f:
    f.write(pdf_bytes)

print('Wrote PDF to', OUTPUT, 'size KB:', os.path.getsize(OUTPUT)/1024)
