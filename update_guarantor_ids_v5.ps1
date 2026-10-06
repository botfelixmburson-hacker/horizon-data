# Script to update guarantor IDs in CSV files using exact Client+Guarantor matching
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv",
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu-arrears-may-jun-jly.csv"
)

Write-Host "Reading guarantor lookup..." -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath

# Re-read Excel files to create proper lookup by Client+Guarantor
$excelFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug disbursement loan balances.xlsx",
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\OCTOBER DISBURSEMENT NYAHURURU (1).xlsx",
    "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep disbursent loan balances.xlsx"
)

$guarantorLookup = @{}

foreach ($excelFile in $excelFiles) {
    Write-Host "Re-reading: $excelFile" -ForegroundColor Green

    try {
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

        for ($row = 2; $row -le $lastRow; $row++) {
            $clientName = $worksheet.Cells.Item($row, 3).Value2
            $guarantorName = $worksheet.Cells.Item($row, 4).Value2
            $guarantorId = $worksheet.Cells.Item($row, 7).Value2

            if (-not $guarantorId) {
                continue
            }

            if ($clientName -and $guarantorName) {
                # Create key: Client Name|Guarantor Name
                $key = "$clientName|$guarantorName"
                if (-not $guarantorLookup.ContainsKey($key)) {
                    $guarantorLookup[$key] = $guarantorId
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

Write-Host "Loaded $($guarantorLookup.Count) guarantor entries" -ForegroundColor Green

foreach ($targetFile in $targetFiles) {
    Write-Host "`nProcessing: $targetFile" -ForegroundColor Green

    $data = Import-Csv -Path $targetFile
    $updatedCount = 0

    foreach ($record in $data) {
        $clientName = $record.'Client Name'
        $guarantorName = $record.'Guarantor Name'
        $currentGuarantorId = $record.'Guarantor ID'

        # Skip if already has guarantor ID
        if ($currentGuarantorId -and $currentGuarantorId -ne "") {
            continue
        }

        # Skip if no guarantor name
        if (-not $guarantorName -or $guarantorName -eq "") {
            continue
        }

        # Look up guarantor ID by Client+Guarantor combination
        $key = "$clientName|$guarantorName"
        if ($guarantorLookup.ContainsKey($key)) {
            $record.'Guarantor ID' = $guarantorLookup[$key]
            $updatedCount++
        }
    }

    Write-Host "Updated $updatedCount records" -ForegroundColor Yellow

    # Export updated data
    $data | Export-Csv -Path $targetFile -NoTypeInformation
    Write-Host "Saved: $targetFile" -ForegroundColor Green
}

Write-Host "`nDone!" -ForegroundColor Green
