# Compare data between Excel source and CSV target
$excelFile = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep disbursent loan balances.xlsx"
$targetFile = "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv"

Write-Host "Reading Excel file data..." -ForegroundColor Green
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($excelFile)
$worksheet = $workbook.Worksheets.Item(1)

$lastRow = 1
for ($row = 2; $row -le 2000; $row++) {
    $clientName = $worksheet.Cells.Item($row, 3).Value2
    if ($clientName) {
        $lastRow = $row
    } else {
        if ($row - $lastRow -gt 10) {
            break
        }
    }
}

$excelClients = @()
for ($row = 2; $row -le $lastRow; $row++) {
    $clientName = $worksheet.Cells.Item($row, 3).Value2
    $guarantorName = $worksheet.Cells.Item($row, 4).Value2
    if ($clientName) {
        $excelClients += "$clientName|$guarantorName"
    }
}

$workbook.Close()
$excel.Quit()

Write-Host "Excel has $($excelClients.Count) records" -ForegroundColor Green

Write-Host "`nReading CSV file data..." -ForegroundColor Green
$csvData = Import-Csv -Path $targetFile
$csvClients = @()
foreach ($record in $csvData) {
    $csvClients += "$($record.'Client Name')|$($record.'Guarantor Name')"
}

Write-Host "CSV has $($csvClients.Count) records" -ForegroundColor Green

Write-Host "`nChecking for matches..." -ForegroundColor Green
$matches = $excelClients | Where-Object { $csvClients -contains $_ }
Write-Host "Found $($matches.Count) exact matches" -ForegroundColor Yellow

if ($matches.Count -gt 0) {
    Write-Host "`nMatches:" -ForegroundColor Green
    $matches | Select-Object -First 10
} else {
    Write-Host "`nNo exact matches found. Checking for client name matches only..." -ForegroundColor Yellow
    $excelClientNames = $excelClients | ForEach-Object { $_ -split '\|' | Select-Object -First 1 }
    $csvClientNames = $csvClients | ForEach-Object { $_ -split '\|' | Select-Object -First 1 }
    $nameMatches = $excelClientNames | Where-Object { $csvClientNames -contains $_ }
    Write-Host "Found $($nameMatches.Count) client name matches" -ForegroundColor Yellow
    if ($nameMatches.Count -gt 0) {
        $nameMatches | Select-Object -First 10
    }
}
