# Check for potential matches
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFile = "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv"

Write-Host "Sample from guarantor lookup:" -ForegroundColor Green
Import-Csv -Path $lookupPath | Select-Object -First 5

Write-Host "`nSample from target file:" -ForegroundColor Green
Import-Csv -Path $targetFile | Select-Object -First 5 'Client Name', 'Guarantor Name', 'Guarantor ID'
