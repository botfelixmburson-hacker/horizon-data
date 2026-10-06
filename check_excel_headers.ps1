# Script to check Excel file headers
$excelFile = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep disbursent loan balances.xlsx"

Write-Host "Reading headers from: $excelFile" -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($excelFile)
    $worksheet = $workbook.Worksheets.Item(1)

    $usedRange = $worksheet.UsedRange
    $colCount = $usedRange.Columns.Count

    Write-Host "Total columns: $colCount" -ForegroundColor Yellow
    Write-Host "Headers:" -ForegroundColor Yellow

    for ($col = 1; $col -le $colCount; $col++) {
        $header = $worksheet.Cells.Item(1, $col).Value2
        Write-Host "Column $($col): $($header)"
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
