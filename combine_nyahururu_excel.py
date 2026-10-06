import pandas as pd
import os

# File paths
files = [
    {"path": r"C:\Users\admin\Desktop\KashLeo data\nyahururu\aug-balances-nyahururu.xlsx", "month": "August"},
    {"path": r"C:\Users\admin\Desktop\KashLeo data\nyahururu\OCTOBER DISBURSEMENT NYAHURURU.xlsx", "month": "October"},
    {"path": r"C:\Users\admin\Desktop\KashLeo data\nyahururu\sep-nyahururu-loan-balances.xlsx", "month": "September"}
]

output_path = r"C:\Users\admin\Desktop\KashLeo data\nyahururu_clients_consolidated.xlsx"

all_data = []

for file_info in files:
    file_path = file_info["path"]
    month = file_info["month"]

    print(f"Processing {month} file...")

    if not os.path.exists(file_path):
        print(f"  File not found: {file_path}")
        continue

    # Read Excel file
    df = pd.read_excel(file_path)

    # Add Month Disbursed column
    df["Month Disbursed"] = month

    all_data.append(df)
    print(f"  Added {len(df)} records")

# Combine all dataframes
print("\nCombining all data...")
combined_df = pd.concat(all_data, ignore_index=True)

# Find the balance column (case-insensitive)
balance_col = None
for col in combined_df.columns:
    if "balance" in col.lower():
        balance_col = col
        break

if balance_col:
    print(f"Found balance column: {balance_col}")

    # Convert balance to numeric, treating non-numeric as 0
    combined_df[balance_col] = pd.to_numeric(combined_df[balance_col], errors='coerce').fillna(0)

    # Sort by balance (descending) - 0 balances will be at bottom
    print("Sorting by Total Loan Balance...")
    combined_df = combined_df.sort_values(by=balance_col, ascending=False)
else:
    print("Warning: Could not find balance column")

print(f"Total records: {len(combined_df)}")

# Export to Excel
print(f"\nExporting to {output_path}...")
combined_df.to_excel(output_path, index=False)

print("Done!")
