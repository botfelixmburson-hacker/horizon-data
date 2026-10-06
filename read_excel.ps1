# PowerShell script to read Excel files
$ErrorActionPreference = "Stop"

$folderPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu"
$files = @{
    "October" = "OCTOBER DISBURSEMENT NYAHURURU.xlsx"
    "August" = "aug-balances-nyahururu.xlsx"
    "September" = "sep-nyahururu-loan-balances.xlsx"
}

# Try to use Excel COM object
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($month in $files.Keys) {
        $filename = $files[$month]
        $filePath = Join-Path $folderPath $filename

        Write-Host "Reading $month file: $filename"
        if (Test-Path $filePath) {
            $workbook = $excel.Workbooks.Open($filePath)
            $worksheet = $workbook.Sheets.Item(1)

            # Get used range
            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count

            Write-Host "Rows: $rowCount, Columns: $colCount"

            # Get column headers (first row)
            $headers = @()
            for ($col = 1; $col -le $colCount; $col++) {
                $headers += $worksheet.Cells.Item(1, $col).Value2
            }
            Write-Host "Headers: $($headers -join ', ')"

            # Get first 5 data rows
            Write-Host "First 5 rows:"
            for ($row = 2; $row -le [Math]::Min(6, $rowCount); $row++) {
                $rowData = @()
                for ($col = 1; $col -le $colCount; $col++) {
                    $rowData += $worksheet.Cells.Item($row, $col).Value2
                }
                Write-Host "$($rowData -join ', ')"
            }

            $workbook.Close($false)
        } else {
            Write-Host "File not found: $filePath"
        }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    Write-Host "Error: $_"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
