# Script to check Excel file data
$excelFile = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep disbursent loan balances.xlsx"

Write-Host "Reading data from: $excelFile" -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($excelFile)
    $worksheet = $workbook.Worksheets.Item(1)

    $usedRange = $worksheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Total rows: $rowCount" -ForegroundColor Yellow
    Write-Host "Total columns: $colCount" -ForegroundColor Yellow
    Write-Host "`nFirst 5 rows of data:" -ForegroundColor Yellow

    for ($row = 1; $row -le [Math]::Min(5, $rowCount); $row++) {
        Write-Host "`nRow $($row):" -ForegroundColor Cyan
        for ($col = 1; $col -le [Math]::Min(20, $colCount); $col++) {
            $value = $worksheet.Cells.Item($row, $col).Value2
            if ($value) {
                Write-Host "  Col $($col): $($value)"
            }
        }
    }

    $workbook.Close()
    $excel.Quit()
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
} finally {
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
