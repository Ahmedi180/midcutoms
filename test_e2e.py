import sys
sys.path.insert(0, '.')
from app import read_uploaded_file, validate_dataframe, process_dataframe

CSV_PATH = r'DREReport - Custom - 2026-05-22T021524.279.csv'

print('=== End-to-End CSV Test ===')
df = read_uploaded_file(CSV_PATH, CSV_PATH)
print('Loaded shape:', df.shape)
print('Columns:', list(df.columns))

errors = validate_dataframe(df)
if errors:
    print('VALIDATION ERRORS:', errors)
    sys.exit(1)

print('Validation: OK')
print()

cleaned_df, stats = process_dataframe(df)
print('Stats:', stats)
print()

print('=== Sample cleaned rows ===')
for i, row in cleaned_df.head(8).iterrows():
    print('Row', i+1, ':', row['Manifested Description'][:80])
