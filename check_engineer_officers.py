import pandas as pd

file = r"C:\Users\admin\Desktop\KashLeo data\draft\ENGINEER BRANCH DISBURSEMENT & OLB Report (1) (3) (1).xlsx"

print("Checking Engineer branch file for loan officer details...")
df = pd.read_excel(file, sheet_name="Sheet1")
print(f"\nColumns: {df.columns.tolist()}")
print(f"\nFirst 5 rows:")
print(df.head())

# Check for loan officer column
if 'LOAN OFFICER' in df.columns:
    print(f"\nUnique loan officers: {df['LOAN OFFICER'].unique()}")
