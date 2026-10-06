# PowerShell script to combine Nyahururu loan data - optimized version
$ErrorActionPreference = "Stop"

$folderPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\final data\nyahururu_clients_combined.csv"

$files = @{
    "October" = "OCTOBER DISBURSEMENT NYAHURURU.xlsx"
    "August" = "aug-balances-nyahururu.xlsx"
    "September" = "sep-nyahururu-loan-balances.xlsx"
}

$allData = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($month in $files.Keys) {
        $filename = $files[$month]
        $filePath = Join-Path $folderPath $filename

        Write-Host "Processing $month file: $filename"

        if (Test-Path $filePath) {
            $workbook = $excel.Workbooks.Open($filePath)
            $worksheet = $workbook.Sheets.Item(1)

            # Find the last row with actual data (used range end)
            $lastRow = $worksheet.UsedRange.Row + $worksheet.UsedRange.Rows.Count - 1
            $lastCol = $worksheet.UsedRange.Column + $worksheet.UsedRange.Columns.Count - 1

            Write-Host "  Data rows: $lastRow, Columns: $lastCol"

            # Get column headers (first row)
            $headers = @()
            for ($col = 1; $col -le $lastCol; $col++) {
                $headerValue = $worksheet.Cells.Item(1, $col).Value2
                if ($headerValue) {
                    $headers += $headerValue.Trim()
                } else {
                    $headers += "Column$col"
                }
            }

            # Find the column index for TOTAL LOAN BALANCE
            $balanceColIndex = 0
            for ($col = 1; $col -le $lastCol; $col++) {
                if ($headers[$col-1] -eq "TOTAL LOAN BALANCE") {
                    $balanceColIndex = $col
                    break
                }
            }

            Write-Host "  TOTAL LOAN BALANCE column index: $balanceColIndex"

            # Read data rows (skip header row)
            $count = 0
            for ($row = 2; $row -le $lastRow; $row++) {
                # Check if first column has data (skip empty rows)
                $firstCell = $worksheet.Cells.Item($row, 1).Value2
                if ($firstCell -eq $null -or $firstCell -eq "") {
                    continue
                }

                $rowData = @()
                for ($col = 1; $col -le $lastCol; $col++) {
                    $cellValue = $worksheet.Cells.Item($row, $col).Value2
                    $rowData += $cellValue
                }

                # Get balance value
                $balance = 0
                if ($balanceColIndex -gt 0) {
                    $balance = $worksheet.Cells.Item($row, $balanceColIndex).Value2
                    if ($balance -eq $null) { $balance = 0 }
                }

                # Only include if balance > 0
                if ($balance -gt 0) {
                    # Add month as first column
                    $rowWithMonth = @($month) + $rowData
                    $allData += $rowWithMonth
                    $count++
                }
            }

            $workbook.Close($false)
            Write-Host "  Added $count records with balance > 0"
        } else {
            Write-Host "  File not found: $filePath"
        }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    # Add "Month" to headers
    $finalHeaders = @("Month Disbursed") + $headers

    # Sort by TOTAL LOAN BALANCE (descending)
    $balanceColIndexFinal = [Array]::IndexOf($finalHeaders, "TOTAL LOAN BALANCE")
    if ($balanceColIndexFinal -ge 0) {
        $allData = $allData | Sort-Object { [double]$($_[$balanceColIndexFinal]) } -Descending
    }

    Write-Host "`nTotal records with balance > 0: $($allData.Count)"

    # Write to CSV
    $csvContent = $finalHeaders -join ","
    foreach ($row in $allData) {
        $csvRow = ($row | ForEach-Object {
            if ($_ -is [string]) {
                '"{0}"' -f ($_ -replace '"', '""')
            } else {
                $_
            }
        }) -join ","
        $csvContent += "`n" + $csvRow
    }

    $csvContent | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Host "`nOutput saved to: $outputPath"

} catch {
    Write-Host "Error: $_"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
