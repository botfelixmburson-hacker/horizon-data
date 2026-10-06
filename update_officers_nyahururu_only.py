import pandas as pd
from datetime import datetime

print("Updating Loan Officer names to Nyahururu branch only...")
print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

# Read files
loans_file = r"C:\Users\admin\Desktop\KashLeo data\complete\FINAL_NYAHURURU_LOANS_WITH_PRODUCT.xlsx"
employees_file = r"C:\Users\admin\Desktop\KashLeo data\complete\kashleo-employees.csv"

print(f"\nReading loans file: {loans_file}")
df_loans = pd.read_excel(loans_file)
print(f"  Records: {len(df_loans)}")

print(f"\nReading employees file: {employees_file}")
df_employees = pd.read_csv(employees_file)
print(f"  Records: {len(df_employees)}")

# Filter only Nyahururu branch loan officers
nyahururu_officers = df_employees[
    (df_employees['Branch'] == 'Nyahururu') &
    (df_employees['Position'] == 'Loan Officer')
]
print(f"\nNyahururu branch loan officers: {len(nyahururu_officers)}")
print(nyahururu_officers[['Name']])

# Create lookup by first name
employee_lookup = {}
for idx, row in nyahururu_officers.iterrows():
    full_name = str(row['Name']).strip()
    first_name = full_name.split()[0].lower()
    employee_lookup[first_name] = full_name

print(f"\nLookup dictionary: {employee_lookup}")

# Update Loan Officer names
updated_count = 0
not_found = []

for idx, row in df_loans.iterrows():
    officer = str(row['Loan Officer']).strip() if pd.notna(row['Loan Officer']) else ""

    if officer and officer != "" and officer != "nan":
        officer_lower = officer.lower()
        if officer_lower in employee_lookup:
            df_loans.at[idx, 'Loan Officer'] = employee_lookup[officer_lower]
            updated_count += 1
        else:
            not_found.append(officer)

print(f"\nUpdated {updated_count} Loan Officer names")
print(f"Loan officers not found in Nyahururu branch: {len(set(not_found))}")

if len(set(not_found)) > 0:
    print(f"\nLoan officers not found:")
    for officer in set(not_found):
        print(f"  - {officer}")

# Save to file
print(f"\nSaving to: {loans_file}")
df_loans.to_excel(loans_file, index=False)

print("\nDone!")
print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
