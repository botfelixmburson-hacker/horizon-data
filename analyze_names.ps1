# Analyze name differences
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFile = "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv"

Write-Host "Guarantor names in lookup:" -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath
$lookupData | Select-Object -ExpandProperty GuarantorName | Sort-Object

Write-Host "`nGuarantor names in target file (first 20 without IDs):" -ForegroundColor Green
$targetData = Import-Csv -Path $targetFile
$targetData | Where-Object { -not $_.'Guarantor ID' -or $_.'Guarantor ID' -eq "" } | Select-Object -First 20 -ExpandProperty 'Guarantor Name'
