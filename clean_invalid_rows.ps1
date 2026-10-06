# Remove invalid rows from nyahururu_branch_loan_balances.csv
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

    # Parse the line
    $values = $line -split ","
    
    # Check if this is a valid data row (should have 17 columns)
    if ($values.Count -eq 17) {
        # Check if the first field (Month Disbursed) is valid
        $month = $values[0].Trim('"')
        if ($month -in @("August", "September", "October", "")) {
            $validLines += $line
        } else {
            Write-Host "  Removed invalid row (invalid month): $line"
            $removedCount++
        }
    } else {
        Write-Host "  Removed invalid row (wrong column count): $line"
        $removedCount++
    }
}

Write-Host "`nRemoved $removedCount invalid rows"
Write-Host "Kept $($validLines.Count) valid rows"

# Write cleaned file
$validLines -join "`n" | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
