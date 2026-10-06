# Convert May, June, July Excel files to CSV
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$files = @(
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\May Disbursement-nyahururu.xls"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\may-disbursement-nyahururu.csv"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\JUNE DISBURSEMENT-nyahururu.xlsx"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\june-disbursement-nyahururu.csv"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\JULY DISBURSEMENT-nyahururu.xlsx"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\july-disbursement-nyahururu.csv"}
)

try {
    foreach ($file in $files) {
        Write-Host "Converting $($file.Path)..."
        if (-not (Test-Path $file.Path)) {
            Write-Host "  File not found: $($file.Path)"
            continue
        }
        $workbook = $excel.Workbooks.Open($file.Path)
        $csvPath = $file.Output
        $workbook.SaveAs($csvPath, 6)  # 6 = xlCSV
        $workbook.Close($false)
        Write-Host "  Saved to $csvPath"
    }
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Conversion complete!"
