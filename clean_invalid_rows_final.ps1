# Manually remove invalid rows from nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading file..."
$lines = Get-Content -Path $nyahururuPath

$validLines = @()
$removedCount = 0

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]

    # Skip empty lines
    if ($line.Trim() -eq "") {
        continue
    }

    # Keep header (first line)
    if ($i -eq 0) {
        $validLines += $line
        continue
    }

    # Check if line starts with valid month in quotes
    if ($line -match '^"(August|September|October)"') {
        $validLines += $line
    } elseif ($line -match '^"","') {
        # Line with empty month but has client name - check if it's valid
        $values = $line -split ","
        if ($values.Count -ge 3 -and $values[2].Trim('"') -ne "") {
            $validLines += $line
        } else {
            Write-Host "  Removed invalid row: $line"
            $removedCount++
        }
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
