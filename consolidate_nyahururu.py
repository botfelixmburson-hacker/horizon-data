import pandas as pd
import glob
import os

# Initialize a list to store all client data
all_clients = []

# Process the reconciled file (most complete)
try:
    df_reconciled = pd.read_csv('final data/nyahururu_branch_loan_balances_reconciled.csv')
    for _, row in df_reconciled.iterrows():
        all_clients.append({
            'Client Name': row.get('Client Name', ''),
            'Client ID': row.get('Client ID', ''),
            'Client Phone': row.get('Client Phone', ''),
            'Guarantor Name': row.get('Guarantor Name', ''),
            'Guarantor ID': row.get('Guarantor ID', ''),
            'Guarantor Phone': row.get('Guarantor Phone', ''),
            'Loan Officer': row.get('Loan Officer', ''),
            'Branch': 'Nyahururu',
            'Disbursement Date': row.get('Disbursement Date', ''),
            'Product Name': row.get('Product Name', ''),
            'Principal': row.get('Principal', ''),
            'Balance': row.get('Balance', ''),
            'Interest Rate': row.get('Interest Rate', ''),
            'Status': row.get('Loan Status', ''),
            'Source File': row.get('Source File', '')
        })
    print(f'Processed reconciled file: {len(df_reconciled)} records')
except Exception as e:
    print(f'Error processing reconciled file: {e}')

# Process July disbursement
try:
    df_july = pd.read_csv('JULY DISBURSEMENT-nyahururu-branch.csv')
    for _, row in df_july.iterrows():
        client_name = row.get('CLIENT\'S NAME', '')
        if pd.notna(client_name) and client_name != '':
            all_clients.append({
                'Client Name': client_name,
                'Client ID': row.get('795493127', ''),
                'Client Phone': row.get('CLIENT\'S CONTACT', ''),
                'Guarantor Name': row.get('GUARANTOR\'S NAME', ''),
                'Guarantor ID': '',
                'Guarantor Phone': '',
                'Loan Officer': row.get('L.O', ''),
                'Branch': 'Nyahururu',
                'Disbursement Date': row.get('Disbursement Date', ''),
                'Product Name': row.get('repayment period', ''),
                'Principal': row.get('AMOUNT DISBURSED', ''),
                'Balance': row.get('TOTAL LOAN BALANCE', ''),
                'Interest Rate': '',
                'Status': 'Active',
                'Source File': 'JULY DISBURSEMENT-nyahururu-branch.csv'
            })
    print(f'Processed July file: {len(df_july)} records')
except Exception as e:
    print(f'Error processing July file: {e}')

# Process August disbursement
try:
    df_august = pd.read_csv('AUGUST DISBURSEMENT NYAHURURU (2).csv')
    # August file has different column structure
    august_cols = df_august.columns.tolist()
    print(f'August columns: {august_cols[:5]}')
    for _, row in df_august.iterrows():
        client_name = row.get(august_cols[1] if len(august_cols) > 1 else '', '')
        if pd.notna(client_name) and client_name != '':
            all_clients.append({
                'Client Name': client_name,
                'Client ID': row.get(august_cols[3] if len(august_cols) > 3 else '', ''),
                'Client Phone': row.get(august_cols[3] if len(august_cols) > 3 else '', ''),
                'Guarantor Name': row.get(august_cols[2] if len(august_cols) > 2 else '', ''),
                'Guarantor ID': '',
                'Guarantor Phone': '',
                'Loan Officer': row.get(august_cols[-1] if len(august_cols) > 0 else '', ''),
                'Branch': 'Nyahururu',
                'Disbursement Date': row.get(august_cols[0] if len(august_cols) > 0 else '', ''),
                'Product Name': row.get(august_cols[5] if len(august_cols) > 5 else '', ''),
                'Principal': row.get(august_cols[4] if len(august_cols) > 4 else '', ''),
                'Balance': row.get(august_cols[10] if len(august_cols) > 10 else '', ''),
                'Interest Rate': '',
                'Status': 'Active',
                'Source File': 'AUGUST DISBURSEMENT NYAHURURU (2).csv'
            })
    print(f'Processed August file: {len(df_august)} records')
except Exception as e:
    print(f'Error processing August file: {e}')

# Process May disbursement
try:
    df_may = pd.read_csv('May Disbursement-nyahururu-branch.csv')
    for _, row in df_may.iterrows():
        client_name = row.get('CLIENT\'S NAME', '')
        if pd.notna(client_name) and client_name != '':
            all_clients.append({
                'Client Name': client_name,
                'Client ID': '',
                'Client Phone': row.get('CLIENT\'S CONTACT', ''),
                'Guarantor Name': row.get('GUARANTOR\'S NAME', ''),
                'Guarantor ID': '',
                'Guarantor Phone': row.get('GUARANTOR\'S CONTACT', ''),
                'Loan Officer': row.get('L.O', ''),
                'Branch': 'Nyahururu',
                'Disbursement Date': row.get('Disbursement Date', ''),
                'Product Name': row.get('repayment period', ''),
                'Principal': row.get('AMOUNT DISBURSED', ''),
                'Balance': row.get('TOTAL LOAN BALANCE', ''),
                'Interest Rate': '',
                'Status': 'Active',
                'Source File': 'May Disbursement-nyahururu-branch.csv'
            })
    print(f'Processed May file: {len(df_may)} records')
except Exception as e:
    print(f'Error processing May file: {e}')

# Process October disbursement
try:
    df_october = pd.read_csv('OCTOBER DISBURSEMENT NYAHURURU.csv')
    for _, row in df_october.iterrows():
        client_name = row.get('CLIENT\'S NAME', '')
        if pd.notna(client_name) and client_name != '':
            all_clients.append({
                'Client Name': client_name,
                'Client ID': row.get('CLIENTS ID', ''),
                'Client Phone': row.get('CLIENT\'S CONTACT', ''),
                'Guarantor Name': row.get('GUARANTOR\'S NAME', ''),
                'Guarantor ID': row.get('GUARANTORS ID', ''),
                'Guarantor Phone': row.get('GUARONTOR CONTACT', ''),
                'Loan Officer': row.get('L.O', ''),
                'Branch': 'Nyahururu',
                'Disbursement Date': row.get('DISBURSEMENT DATE', ''),
                'Product Name': row.get('REPAYMENT PERIOD', ''),
                'Principal': row.get('AMOUNT DISBURSED', ''),
                'Balance': row.get('TOTAL LOAN BALANCE', ''),
                'Interest Rate': '',
                'Status': 'Active',
                'Source File': 'OCTOBER DISBURSEMENT NYAHURURU.csv'
            })
    print(f'Processed October file: {len(df_october)} records')
except Exception as e:
    print(f'Error processing October file: {e}')

# Create DataFrame and save
df_all = pd.DataFrame(all_clients)
df_all.to_csv('nyahururu_clients_consolidated.csv', index=False)
print(f'Total records consolidated: {len(df_all)}')
print('File saved: nyahururu_clients_consolidated.csv')
