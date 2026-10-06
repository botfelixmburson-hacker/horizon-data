# Manually remove invalid rows from nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading file..."
$lines = Get-Content -Path $nyahururuPath

$validLines = @()
$removedCount = 0

foreach ($line in $lines) {
    # Skip empty lines
    if ($line.Trim() -eq "") {
        continue
    }

    # Check if line starts with valid month in quotes
    if ($line -match '^"(August|September|October)"') {
        $validLines += $line
    } elseif ($line -match '^"","') {
        # Line with empty month but has client name - keep it
        $validLines += $line
    } else {
        Write-Host "  Removed invalid row: $line"
        $removedCount++
    }
}

Write-Host "`nRemoved $removedCount invalid rows"
Write-Host "Kept $($validLines.Count) valid rows"

# Write cleaned file
$validLines -join "`n" | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
