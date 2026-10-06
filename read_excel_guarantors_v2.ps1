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

        # Fixed column indices based on observed structure
        $clientNameCol = 3  # CLIENT'S NAME
        $guarantorNameCol = 4  # GUARANTOR'S NAME
        $guarantorIdCol = 7  # Guarantor ID (appears in column 7)

        Write-Host "Reading rows 2 to $rowCount" -ForegroundColor Yellow

        for ($row = 2; $row -le $rowCount; $row++) {
            $clientName = $worksheet.Cells.Item($row, $clientNameCol).Value2
            $guarantorName = $worksheet.Cells.Item($row, $guarantorNameCol).Value2
            $guarantorId = $worksheet.Cells.Item($row, $guarantorIdCol).Value2

            # Skip if guarantor ID is empty or null
            if (-not $guarantorId) {
                continue
            }

            if ($clientName -and $guarantorName) {
                $key = "$clientName|$guarantorName"
                if (-not $guarantorData.ContainsKey($key)) {
                    $guarantorData[$key] = $guarantorId
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
    $parts = $_.Key -split '\|'
    [PSCustomObject]@{
        ClientName = $parts[0]
        GuarantorName = $parts[1]
        GuarantorID = $_.Value
    }
} | Export-Csv -Path "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -NoTypeInformation

Write-Host "Guarantor lookup saved to: C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -ForegroundColor Green
