# Combine Nyahururu Excel files directly
$ErrorActionPreference = "Stop"

$excelFiles = @(
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug-balances-nyahururu.xlsx"; Month = "August"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\OCTOBER DISBURSEMENT NYAHURURU.xlsx"; Month = "October"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep-nyahururu-loan-balances.xlsx"; Month = "September"}
)

$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu_clients_consolidated.xlsx"

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$allData = @()

try {
    foreach ($fileInfo in $excelFiles) {
        Write-Host "Processing $($fileInfo.Month) file..."
        $filePath = $fileInfo.Path

        if (-not (Test-Path $filePath)) {
            Write-Host "  File not found: $filePath"
            continue
        }

        $workbook = $excel.Workbooks.Open($filePath)
        $worksheet = $workbook.Sheets.Item(1)

        # Find the last row and column
        $lastRow = $worksheet.UsedRange.Rows.Count
        $lastCol = $worksheet.UsedRange.Columns.Count

        # Get headers from first row
        $headers = @()
        for ($col = 1; $col -le $lastCol; $col++) {
            $headers += $worksheet.Cells.Item(1, $col).Value2
        }

        # Find the column index for TOTAL LOAN BALANCE
        $balanceColIndex = -1
        for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -like "*BALANCE*" -or $headers[$i] -like "*balance*") {
                $balanceColIndex = $i + 1
                break
            }
        }

        if ($balanceColIndex -eq -1) {
            Write-Host "  Warning: Could not find balance column, using last column"
            $balanceColIndex = $lastCol
        }

        # Read data rows (skip header row)
        $count = 0
        for ($row = 2; $row -le $lastRow; $row++) {
            $rowData = @{}
            for ($col = 1; $col -le $lastCol; $col++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Value2
                $headerName = if ($headers[$col - 1]) { $headers[$col - 1] } else { "Column$col" }
                $rowData[$headerName] = $cellValue
            }

            # Include all records (including 0 balance)
            $rowData["Month Disbursed"] = $fileInfo.Month
            $allData += [PSCustomObject]$rowData
            $count++
        }

        Write-Host "  Added $count records"

        $workbook.Close($false)
    }

    # Sort by Total Loan Balance (descending) - 0 balances will be at bottom
    Write-Host "`nSorting by Total Loan Balance..."
    $allData = $allData | Sort-Object {
        $balanceCol = $_.PSObject.Properties | Where-Object { $_.Name -like "*BALANCE*" -or $_.Name -like "*balance*" } | Select-Object -First 1
        if ($balanceCol) {
            $val = $_.($balanceCol.Name)
            if ($val -is [double] -or $val -is [int]) {
                $val
            } elseif ($val -match "^\d+\.?\d*$") {
                [double]$val
            } else {
                0
            }
        } else {
            0
        }
    } -Descending

    Write-Host "Total records: $($allData.Count)"

    # Export to Excel
    Write-Host "Exporting to Excel..."
    $newWorkbook = $excel.Workbooks.Add()
    $newWorksheet = $newWorkbook.Sheets.Item(1)

    # Write headers
    $allProperties = $allData[0].PSObject.Properties.Name
    for ($col = 0; $col -lt $allProperties.Count; $col++) {
        $newWorksheet.Cells.Item(1, $col + 1).Value2 = $allProperties[$col]
    }

    # Write data
    for ($row = 0; $row -lt $allData.Count; $row++) {
        for ($col = 0; $col -lt $allProperties.Count; $col++) {
            $propName = $allProperties[$col]
            $value = $allData[$row].$propName
            if ($null -ne $value) {
                $newWorksheet.Cells.Item($row + 2, $col + 1).Value2 = $value
            }
        }
    }

    # Save and close
    $newWorkbook.SaveAs($outputPath)
    $newWorkbook.Close($false)

    Write-Host "`nOutput saved to: $outputPath"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
