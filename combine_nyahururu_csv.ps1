# Combine Nyahururu CSV files - filter by balance > 0 and sort
$ErrorActionPreference = "Stop"

$csvFolder = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\final data\nyahururu_clients_combined.csv"

$files = @{
    "October" = "October.csv"
    "August" = "August.csv"
    "September" = "September.csv"
}

$allData = @()

foreach ($month in $files.Keys) {
    $filename = $files[$month]
    $filePath = Join-Path $csvFolder $filename

    Write-Host "Processing $month file: $filename"

    if (Test-Path $filePath) {
        $lines = Get-Content $filePath

        # Get headers from first line
        $headers = $lines[0].Split(",")

        # Find column index for TOTAL LOAN BALANCE or balances
        $balanceColIndex = -1
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i].Trim() -eq "TOTAL LOAN BALANCE" -or $headers[$i].Trim() -eq "balances") {
                $balanceColIndex = $i
                break
            }
        }

        Write-Host "  TOTAL LOAN BALANCE column index: $balanceColIndex"

        $count = 0
        # Process data rows (skip header)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            if ($line -eq "") { continue }

            $values = $line.Split(",")

            # Get balance value
            $balance = 0
            if ($balanceColIndex -ge 0 -and $balanceColIndex -lt $values.Count) {
                $balanceStr = $values[$balanceColIndex].Trim()
                if ($balanceStr -ne "") {
                    $balance = [double]$balanceStr
                }
            }

            # Only include if balance > 0
            if ($balance -gt 0) {
                # Add month as first column
                $rowWithMonth = @($month) + $values
                $allData += $rowWithMonth
                $count++
            }
        }

        Write-Host "  Added $count records with balance > 0"
    }
}

# Add "Month Disbursed" to headers
$finalHeaders = @("Month Disbursed") + $headers

# Sort by TOTAL LOAN BALANCE or balances (descending)
$balanceColIndexFinal = [Array]::IndexOf($finalHeaders, "TOTAL LOAN BALANCE")
if ($balanceColIndexFinal -lt 0) {
    $balanceColIndexFinal = [Array]::IndexOf($finalHeaders, "balances")
}
if ($balanceColIndexFinal -ge 0) {
    $allData = $allData | Sort-Object { [double]$($_[$balanceColIndexFinal]) } -Descending
}

Write-Host "`nTotal records with balance > 0: $($allData.Count)"

# Write to CSV
$csvContent = $finalHeaders -join ","
foreach ($row in $allData) {
    $csvRow = ($row | ForEach-Object {
        $val = $_.Trim()
        if ($val -match ',') {
            '"{0}"' -f ($val -replace '"', '""')
        } else {
            $val
        }
    }) -join ","
    $csvContent += "`n" + $csvRow
}

$csvContent | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "`nOutput saved to: $outputPath"
