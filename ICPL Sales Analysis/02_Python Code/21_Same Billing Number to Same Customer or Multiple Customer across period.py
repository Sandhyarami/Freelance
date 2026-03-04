import pandas as pd

# 📌 Input file path
file_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/01_Input Data/02_Clean Data/Sales Data Clean.xlsx"

# 🔹 Load 'Working' sheet only ONCE
df = pd.read_excel(file_path, sheet_name='Working')

# ✅ Group by Voucher No.
grouped = df.groupby('Voucher No.').agg({
    'Date': lambda x: '|'.join(sorted(set(map(str, x)))),
    'Customer': lambda x: '|'.join(sorted(set(map(str, x))))
}).reset_index()

# ✅ Count unique Dates & Customers
grouped['Date_Count'] = grouped['Date'].apply(lambda x: len(x.split('|')))
grouped['Customer_Count'] = grouped['Customer'].apply(lambda x: len(x.split('|')))

# ✅ Remark logic
def get_remark(row):
    if row['Date_Count'] > 1 and row['Customer_Count'] > 1:
        return 'Multiple Dates & Multiple Customers'
    elif row['Date_Count'] > 1:
        return 'Multiple Dates'
    elif row['Customer_Count'] > 1:
        return 'Multiple Customers'
    else:
        return ''

grouped['Remark'] = grouped.apply(get_remark, axis=1)

# ✅ Filter 1️⃣: Only Multiple Dates — only Date related columns
df_multiple_dates = grouped[grouped['Date_Count'] > 1][['Voucher No.', 'Date', 'Date_Count', 'Remark']]
df_multiple_dates.insert(0, 'SR No.', range(1, len(df_multiple_dates) + 1))

# ✅ Filter 2️⃣: Only Multiple Customers — only Customer related columns
df_multiple_customers = grouped[grouped['Customer_Count'] > 1][['Voucher No.', 'Customer', 'Customer_Count', 'Remark']]
df_multiple_customers.insert(0, 'SR No.', range(1, len(df_multiple_customers) + 1))

# 🔹 Save both to single output file with 2 sheets
output_path = r"D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/02_Output Data/21_Same Billing Number to Same Customer or Multiple Customer across period.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_multiple_dates.to_excel(writer, sheet_name='Multiple_Dates', index=False)
    df_multiple_customers.to_excel(writer, sheet_name='Multiple_Customers', index=False)

print(f"✅ DONE! Check your output: {output_path}")
