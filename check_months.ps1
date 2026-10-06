# Check months in different files
$targetFile = "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv"

Write-Host "Months in target CSV:" -ForegroundColor Green
$data = Import-Csv -Path $targetFile
$data.'Month Disbursed' | Group-Object | Select-Object Name, Count

Write-Host "`nTotal records per month:" -ForegroundColor Green
$data | Group-Object 'Month Disbursed' | Select-Object Name, Count
