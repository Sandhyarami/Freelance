import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime

# 🔹 Load Excel file
input_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/01_Input Data/02_Clean Data/Sales Data Clean.xlsx"
sheet_name = 'Working'

# 🔹 Read Excel into DataFrame
data = pd.read_excel(input_path, sheet_name=sheet_name, engine='openpyxl')

# 🔹 Filter only required columns
required_columns = ['Date', 'Customer', 'Item Name', 'Voucher Type', 'Voucher No.', 'Sales Type']
filtered_data = data[required_columns]

# 🔹 Keep only Export Sales
export_data = filtered_data[filtered_data['Sales Type'].astype(str).str.strip().str.lower() == 'export'].copy()

# 🔹 Drop Voucher No. used on multiple different Dates
voucher_date_counts = export_data.groupby('Voucher No.')['Date'].nunique().reset_index()
invalid_vouchers = voucher_date_counts[voucher_date_counts['Date'] > 1]['Voucher No.']
export_data = export_data[~export_data['Voucher No.'].isin(invalid_vouchers)].copy()

# 🔹 Remove duplicate records based on Date + Voucher No.
export_data_unique = export_data.drop_duplicates(subset=['Date', 'Voucher No.']).copy()

# 🔹 Convert Date to datetime
export_data_unique['Date'] = pd.to_datetime(export_data_unique['Date'], errors='coerce')

# 🔹 Extract numeric part of Voucher No.
export_data_unique['Voucher Numeric'] = export_data_unique['Voucher No.'].astype(str).str.split('/').str[0]
export_data_unique['Voucher Numeric'] = pd.to_numeric(export_data_unique['Voucher Numeric'], errors='coerce')

# 🔹 Filter Date > 01-Apr-2024
export_data_unique = export_data_unique[export_data_unique['Date'] > pd.to_datetime('2024-04-01')]

# 🔹 Sort by Date and Voucher No.
export_data_unique = export_data_unique.sort_values(by=['Date', 'Voucher Numeric']).reset_index(drop=True)

# 🔹 Create shift columns
export_data_unique['Prev Voucher No.'] = export_data_unique['Voucher Numeric'].shift(1)
export_data_unique['Prev Date'] = export_data_unique['Date'].shift(1)

# 🔹 Calculate difference only if Date is different from previous row
export_data_unique['Difference'] = np.where(
    export_data_unique['Date'] != export_data_unique['Prev Date'],
    export_data_unique['Voucher Numeric'] - export_data_unique['Prev Voucher No.'],
    np.nan  # Don't calculate if date is same
)

# 🔹 Keep only rows where Difference is NOT 1 (i.e., series break across different dates)
non_sequential = export_data_unique[export_data_unique['Difference'].notna() & (export_data_unique['Difference'] != 1)].copy()

export_data_unique['Voucher Numeric'] = pd.to_numeric(export_data_unique['Voucher Numeric'], errors='coerce')
export_data_unique.rename(columns={'Voucher Numeric': 'Current Voucher No.'}, inplace=True)

# 🔹 Add Sr. No.
non_sequential.insert(0, 'Sr. No.', range(1, len(non_sequential) + 1))

# 🔹 Save to Excel
output_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/02_Output Data/19_Invoice Series discrepancies.xlsx"    
non_sequential.to_excel(output_path, index=False)

print(f"✅ Invoice series discrepancies saved to: {output_path}")
