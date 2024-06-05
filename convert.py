import pandas as pd

# Path to the Excel file in the repository
excel_file = 'AIESEC_Data.xlsx'  # Update this path to your actual Excel file

# Load the Excel file
df = pd.read_excel(excel_file)

# Save as CSV
csv_file = 'data.csv'
df.to_csv(csv_file, index=False)

print(f'Converted {excel_file} to {csv_file}')
