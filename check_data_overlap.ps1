# Check data overlap between Excel and CSV
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFile = "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv"

Write-Host "Client names in Excel lookup:" -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath
$lookupClients = $lookupData | ForEach-Object { $_.Key -split '\|' | Select-Object -First 1 } | Sort-Object -Unique
$lookupClients | Select-Object -First 10

Write-Host "`nClient names in target CSV:" -ForegroundColor Green
$targetData = Import-Csv -Path $targetFile
$targetClients = $targetData | Select-Object -ExpandProperty 'Client Name' | Sort-Object -Unique
$targetClients | Select-Object -First 10

Write-Host "`nChecking for matching client names..." -ForegroundColor Green
$matches = $lookupClients | Where-Object { $targetClients -contains $_ }
Write-Host "Found $($matches.Count) matching client names out of $($lookupClients.Count) in lookup"
