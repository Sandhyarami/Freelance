import pandas as pd

# Load Sales Data
sales_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/01_Input Data/02_Clean Data/Sales Data Clean.xlsx"
sales_df = pd.read_excel(sales_path, sheet_name='Working', engine='openpyxl')

# Load Credit Note Register
credit_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/01_Input Data/02_Clean Data/Credit Not Register.xlsx"
credit_df = pd.read_excel(credit_path, sheet_name='Credit Note Register', engine='openpyxl')

# Clean column names
sales_df.columns = sales_df.columns.str.strip()
credit_df.columns = credit_df.columns.str.strip()

# Rename columns for clarity before merge
sales_df = sales_df.rename(columns={
    'Date': 'Date',
    'Customer': 'Customer',
    'Item Name': 'Item Name',
    'Voucher No.': 'Voucher_No',
    'Quantity': 'Quantity'
})

credit_df = credit_df.rename(columns={
    'Date': 'Return_Date',
    'Customer': 'Return_Customer',
    'Item Name': 'Return_Item_Name',
    'Voucher Ref. No.': 'Return_Voucher_Ref_No',
    'Quantity': 'Return_Quantity'
})

# Merge only matching records on voucher numbers
merged_df = pd.merge(
    sales_df,
    credit_df,
    left_on='Voucher_No',
    right_on='Return_Voucher_Ref_No',
    how='inner'
)

# ✅ Filter where Return_Item_Name is not blank
merged_df = merged_df[merged_df['Return_Item_Name'].notna() & (merged_df['Return_Item_Name'].astype(str).str.strip() != '')]

# ✅ Keep only rows where Item Name and Return Item Name match exactly
merged_df = merged_df[merged_df['Item Name'].astype(str).str.strip() == merged_df['Return_Item_Name'].astype(str).str.strip()]

# Drop duplicate rows if any
merged_df = merged_df.drop_duplicates()

# Final columns selection
final_df = merged_df[['Date', 'Return_Date', 'Customer', 'Return_Customer', 'Item Name', 'Return_Item_Name', 'Voucher_No', 'Return_Voucher_Ref_No', 'Quantity', 'Return_Quantity']]
final_df.columns = ['Date', 'Return Date', 'Customer', 'Return Customer', 'Item Name', 'Return_Item_Name', 'Voucher No.', 'Return Voucher Ref. No.', 'Quantity', 'Return Quantity']

# ✅ Add Quantity Difference
final_df['Quantity Difference'] = final_df['Quantity'] - final_df['Return Quantity']

# Save to single Excel sheet
output_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/02_Output Data/16_Sales return quantity variance.xlsx"
final_df.to_excel(output_path, index=False, sheet_name=' Quantity Returns')

print(f"✅ Only matched item records saved to: {output_path}")
