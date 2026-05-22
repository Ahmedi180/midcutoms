from app import read_uploaded_file, process_multi_item_row
import os

CSV_PATH = r'DREReport - Custom - 2026-05-22T021524.279.csv'
TARGET_COUNTRY = 'US'

if not os.path.exists(CSV_PATH):
    print('CSV file not found:', CSV_PATH)
    raise SystemExit(1)

# Read file using app helper (handles encodings)
try:
    df = read_uploaded_file(CSV_PATH, os.path.basename(CSV_PATH))
except Exception as e:
    print('Failed to read CSV:', e)
    raise

# Apply country filter if present
if 'Recip Cntry' in df.columns and TARGET_COUNTRY != 'ALL':
    before = len(df)
    df = df[df['Recip Cntry'].astype(str).str.strip().str.upper() == TARGET_COUNTRY.strip().upper()]

# Remove envelopes
if 'Service Type' in df.columns:
    df = df[~df['Service Type'].astype(str).str.contains('Envelope', case=False, na=False)]

total = len(df)
invalid_rows = []
valid_count = 0

for idx, row in df.iterrows():
    manifested = str(row.get('Manifested Description', ''))
    ce_hs_raw = str(row.get('CE Item HSCode', ''))
    cleaned, has_mid, has_hs = process_multi_item_row(manifested, ce_hs_raw)
    if not has_mid or not has_hs:
        tracking = str(row.get('Tracking Number', '')).strip() or f'Row {idx+1}'
        reasons = []
        if not has_mid:
            reasons.append('No valid MID code')
        if not has_hs:
            reasons.append('No HS code found')
        invalid_rows.append({
            'Tracking Number': tracking,
            'Manifested Description': manifested,
            'Reason': ' | '.join(reasons)
        })
    else:
        valid_count += 1

print(f'Total shipments after country/envelope filter: {total}')
print(f'Valid shipments: {valid_count}')
print(f'Invalid shipments (need review): {len(invalid_rows)}')
print()

for r in invalid_rows:
    print(f"- {r['Tracking Number']}: {r['Reason']}")
    print('  Manifested Description:', r['Manifested Description'])
    print()
