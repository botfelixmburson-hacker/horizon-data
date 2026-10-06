$data = Import-Csv 'C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv'
$data | Select-Object -Last 5 | ForEach-Object {
    Write-Host "Month: $($_.'Month Disbursed'), Client: $($_.'CLIENT''S NAME'), Contact: $($_.'CLIENT''S CONTACT')"
}
