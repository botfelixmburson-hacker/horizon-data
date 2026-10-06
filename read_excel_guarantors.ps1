# Script to read guarantor IDs from Excel files
$excelFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug disbursement loan balances.xlsx",
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\OCTOBER DISBURSEMENT NYAHURURU (1).xlsx",
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep disbursent loan balances.xlsx"
)

$guarantorData = @{}

foreach ($excelFile in $excelFiles) {
    Write-Host "Reading: $excelFile" -ForegroundColor Green

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($excelFile)
        $worksheet = $workbook.Worksheets.Item(1)

        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count

        # Find column indices
        $clientNameCol = 0
        $guarantorNameCol = 0
        $guarantorIdCol = 0

        for ($col = 1; $col -le $colCount; $col++) {
            $header = $worksheet.Cells.Item(1, $col).Value2
            if ($header -like "*Client Name*") { $clientNameCol = $col }
            if ($header -like "*Guarantor Name*") { $guarantorNameCol = $col }
            if ($header -like "*Guarantor ID*") { $guarantorIdCol = $col }
        }

        Write-Host "Found columns - Client Name: $clientNameCol, Guarantor Name: $guarantorNameCol, Guarantor ID: $guarantorIdCol" -ForegroundColor Yellow

        if ($clientNameCol -gt 0 -and $guarantorNameCol -gt 0 -and $guarantorIdCol -gt 0) {
            for ($row = 2; $row -le $rowCount; $row++) {
                $clientName = $worksheet.Cells.Item($row, $clientNameCol).Value2
                $guarantorName = $worksheet.Cells.Item($row, $guarantorNameCol).Value2
                $guarantorId = $worksheet.Cells.Item($row, $guarantorIdCol).Value2

                if ($clientName -and $guarantorName -and $guarantorId) {
                    $key = "$clientName|$guarantorName"
                    if (-not $guarantorData.ContainsKey($key)) {
                        $guarantorData[$key] = $guarantorId
                    }
                }
            }
        }

        $workbook.Close()
        $excel.Quit()
    } catch {
        Write-Host "Error reading $($excelFile): $_" -ForegroundColor Red
    } finally {
        if ($excel) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

Write-Host "`nTotal guarantor entries found: $($guarantorData.Count)" -ForegroundColor Green

# Export guarantor data to CSV for reference
$guarantorData.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        Key = $_.Key
        GuarantorID = $_.Value
    }
} | Export-Csv -Path "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -NoTypeInformation

Write-Host "Guarantor lookup saved to: C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -ForegroundColor Green
