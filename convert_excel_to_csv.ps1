# Convert Excel files to CSV
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$files = @(
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug-balances-nyahururu.xlsx"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug-balances-nyahururu.csv"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\OCTOBER DISBURSEMENT NYAHURURU.xlsx"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\october-disbursement-nyahururu.csv"},
    @{Path = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep-nyahururu-loan-balances.xlsx"; Output = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep-nyahururu-loan-balances.csv"}
)

try {
    foreach ($file in $files) {
        Write-Host "Converting $($file.Path)..."
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
