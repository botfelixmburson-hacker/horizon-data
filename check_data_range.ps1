# Check actual data range in Excel files
$ErrorActionPreference = "Stop"

$folderPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu"

$files = @{
    "October" = "OCTOBER DISBURSEMENT NYAHURURU.xlsx"
    "August" = "aug-balances-nyahururu.xlsx"
    "September" = "sep-nyahururu-loan-balances.xlsx"
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($month in $files.Keys) {
        $filename = $files[$month]
        $filePath = Join-Path $folderPath $filename

        Write-Host "`n=== $month ==="
        Write-Host "File: $filename"

        if (Test-Path $filePath) {
            $workbook = $excel.Workbooks.Open($filePath)
            $worksheet = $workbook.Sheets.Item(1)

            # Find last row with data by checking backwards
            $lastRow = $worksheet.UsedRange.Row + $worksheet.UsedRange.Rows.Count - 1
            $lastCol = $worksheet.UsedRange.Column + $worksheet.UsedRange.Columns.Count - 1

            Write-Host "Used range: Row 1 to $lastRow, Col 1 to $lastCol"

            # Search backwards for the last row with actual data
            $actualLastRow = 1
            for ($row = $lastRow; $row -gt 1; $row--) {
                $firstCell = $worksheet.Cells.Item($row, 1).Value2
                if ($firstCell -ne $null -and $firstCell -ne "") {
                    $actualLastRow = $row
                    break
                }
            }

            Write-Host "Actual last row with data: $actualLastRow"

            # Check last 10 rows
            Write-Host "Last 10 rows preview:"
            for ($row = [Math]::Max(2, $actualLastRow - 9); $row -le $actualLastRow; $row++) {
                $name = $worksheet.Cells.Item($row, 3).Value2
                $balance = $worksheet.Cells.Item($row, 16).Value2
                Write-Host "  Row ${row}: ${name} - Balance: ${balance}"
            }

            $workbook.Close($false)
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
