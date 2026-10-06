# Script to read guarantor IDs from Excel files - CORRECTED
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

        # Find the last row with actual data
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

        Write-Host "Found data up to row: $lastRow" -ForegroundColor Yellow

        # Column indices based on actual structure
        $clientNameCol = 3        # CLIENT'S NAME
        $guarantorNameCol = 4     # GUARANTOR'S NAME
        $guarantorIdCol = 7       # Guarantor ID

        for ($row = 2; $row -le $lastRow; $row++) {
            $clientName = $worksheet.Cells.Item($row, $clientNameCol).Value2
            $guarantorName = $worksheet.Cells.Item($row, $guarantorNameCol).Value2
            $guarantorId = $worksheet.Cells.Item($row, $guarantorIdCol).Value2

            if (-not $guarantorId) {
                continue
            }

            if ($clientName -and $guarantorName) {
                # Create key: Guarantor Name only (since we want to match by guarantor name)
                $key = $guarantorName.Trim()
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

# Export guarantor data to CSV
$guarantorData.GetEnumerator() | ForEach-Object {
    [PSCustomObject]@{
        GuarantorName = $_.Key
        GuarantorID = $_.Value
    }
} | Export-Csv -Path "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -NoTypeInformation

Write-Host "Guarantor lookup saved to: C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv" -ForegroundColor Green
