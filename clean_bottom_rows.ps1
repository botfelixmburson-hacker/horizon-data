# Remove corrupted rows at the bottom of the file
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading file..."
$lines = Get-Content -Path $nyahururuPath

Write-Host "Total lines: $($lines.Count)"

# Remove the last 3 lines which appear to be corrupted
$validLines = $lines[0..($lines.Count - 4)]

Write-Host "Kept $($validLines.Count) lines"
Write-Host "Removed 3 corrupted lines from the end"

# Write cleaned file
$validLines -join "`n" | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
