import pandas as pd
from collections import OrderedDict

# === Load Excel file ===
input_file = "D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/01_Input Data/02_Clean Data/Sales Data Clean.xlsx"
df = pd.read_excel(input_file)

# === Clean column names ===
df.columns = df.columns.str.strip()

# === Required columns ===
required_columns = ['Date', 'Customer', 'Item Name', 'Rate', 'Quantity']
missing = [col for col in required_columns if col not in df.columns]
if missing:
    raise ValueError(f"Missing columns: {missing}")

# === Convert 'Date' to datetime ===
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# === Drop rows with missing critical values ===
df = df.dropna(subset=['Date', 'Customer', 'Item Name', 'Rate', 'Quantity'])

# === Normalize strings ===
df['Customer'] = df['Customer'].astype(str).str.strip()
df['Item Name'] = df['Item Name'].astype(str).str.strip()

# === Handle missing 'Voucher Type' ===
if 'Voucher Type' not in df.columns:
    df['Voucher Type'] = ''  # Add blank column if not present

# === Group by Date, Item Name, Voucher Type ===
grouped = df.groupby(['Date', 'Item Name', 'Voucher Type'])

report_records = []

for (date, item, vtype), group in grouped:
    Customers = group['Customer'].unique()
    if len(Customers) <= 1:
        continue  # Skip if only one vendor

    Customer_names = []
    Customer_rates = []
    Customer_qtys = []

    all_rates = []
    all_qtys = []

    for vendor, sub in group.groupby('Customer'):
        Customer_names.append(vendor)

        # Rate with count
        rate_counts = OrderedDict()
        for r in sub['Rate']:
            rate_counts[r] = rate_counts.get(r, 0) + 1
            all_rates.append(r)
        rate_str = ' | '.join([f"{r}({c})" for r, c in rate_counts.items()])
        Customer_rates.append(rate_str)

        # Quantity with count
        qty_counts = OrderedDict()
        for q in sub['Quantity']:
            qty_counts[q] = qty_counts.get(q, 0) + 1
            all_qtys.append(q)
        qty_str = ' | '.join([f"{q}({c})" for q, c in qty_counts.items()])
        Customer_qtys.append(qty_str)

    # === Compute Rate Difference ===
    try:
        rate_diff = max(all_rates) - min(all_rates)
    except:
        rate_diff = 0

    # === Skip if all vendors gave same rate ===
    if rate_diff == 0:
        continue

    # === Quantity comparison ===
    same_qty_flag = len(set(all_qtys)) == 1

    # === Store the final record ===
    report_records.append({
        'Date': date,
        'Item Name': item,
        'Voucher Type': vtype,
        'Customer Count': len(Customers),
        'List of Customers': ' | '.join(Customer_names),
        'Customer Rate List with Count': ' | '.join(Customer_rates),
        'Customer Quantity List with Count': ' | '.join(Customer_qtys),
        'Rate Difference': rate_diff,
        'Same Quantity Flag': 'Yes' if same_qty_flag else 'No'
    })

# === Create final report DataFrame ===
df_report = pd.DataFrame(report_records)

# === Add Sr. No. ===
df_report.insert(0, 'Sr.No.', range(1, len(df_report) + 1))

# === Sort by Rate Difference ascending ===
df_report = df_report.sort_values(by='Rate Difference', ascending=True)

# === Split into two sheets ===
same_qty_df = df_report[df_report['Same Quantity Flag'] == 'Yes'].drop(columns=['Same Quantity Flag'])
diff_qty_df = df_report[df_report['Same Quantity Flag'] == 'No'].drop(columns=['Same Quantity Flag'])

# === Save to Excel with two sheets ===
output_path = "D:/Sandhya/Analysis Insights/Data analysis/ICPL/31-7-2025/02_Output Data/29_Multiple Rates for Same Material (with same quanity) to Different Customer on Same date.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    same_qty_df.to_excel(writer, sheet_name='Same Quantity', index=False)
    diff_qty_df.to_excel(writer, sheet_name='Different Quantity', index=False)

print(f"✅ Report saved to: {output_path}")
