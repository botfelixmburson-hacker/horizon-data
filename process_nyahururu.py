import pandas as pd
import os

# Define the folder path
folder_path = r"C:\Users\admin\Desktop\KashLeo data\nyahururu"

# File names
files = {
    "October": "OCTOBER DISBURSEMENT NYAHURURU.xlsx",
    "August": "aug-balances-nyahururu.xlsx",
    "September": "sep-nyahururu-loan-balances.xlsx"
}

# Read all files
dfs = {}
for month, filename in files.items():
    file_path = os.path.join(folder_path, filename)
    print(f"\nReading {month} file: {filename}")
    df = pd.read_excel(file_path)
    print(f"Columns: {list(df.columns)}")
    print(f"Shape: {df.shape}")
    dfs[month] = df

# Save info for review
for month, df in dfs.items():
    print(f"\n{'='*50}")
    print(f"{month} - First 5 rows:")
    print(df.head())
