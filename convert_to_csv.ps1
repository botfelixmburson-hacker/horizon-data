# Convert Excel files to CSV first
$ErrorActionPreference = "Stop"

$folderPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu"
$csvFolder = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv"

# Create CSV folder if it doesn't exist
if (-not (Test-Path $csvFolder)) {
    New-Item -ItemType Directory -Path $csvFolder | Out-Null
}

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
        $csvPath = Join-Path $csvFolder "${month}.csv"

        Write-Host "Converting $month to CSV..."

        if (Test-Path $filePath) {
            $workbook = $excel.Workbooks.Open($filePath)
            $worksheet = $workbook.Sheets.Item(1)

            # Save as CSV
            $csvFullPath = $excel.Workbooks.Open($filePath)
            $csvFullPath.SaveAs($csvPath, 6)  # 6 = xlCSV
            $csvFullPath.Close($false)

            Write-Host "  Saved to: $csvPath"
        }
    }

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    Write-Host "`nConversion complete!"

} catch {
    Write-Host "Error: $_"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
