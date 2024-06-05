import pandas as pd

# Path to the Excel file in the repository
excel_file = 'AIESEC_Data.xlsx'  # Update this path to your actual Excel file

# Load the Excel file
xls = pd.ExcelFile(excel_file)

# Iterate through each sheet
for sheet_name in xls.sheet_names:
    # Read each sheet into a DataFrame
    df = pd.read_excel(xls, sheet_name)
    
    # Define the CSV file name based on the sheet name
    csv_file = f'{sheet_name}.csv'
    
    # Save the DataFrame as a CSV file
    df.to_csv(csv_file, index=False)
    
    print(f'Converted sheet "{sheet_name}" to {csv_file}')