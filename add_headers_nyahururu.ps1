# Add headers to nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$inputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances_with_headers.csv"

# Define headers based on the reference CSV structure
$headers = @(
    "Month Disbursed",
    "Disbursement Date",
    "Client Name",
    "Client Contact",
    "Client ID",
    "Guarantor Name",
    "Guarantor Contact",
    "Guarantor ID",
    "Amount Disbursed",
    "Repayment Period",
    "Processing Fees",
    "Interest",
    "OLB",
    "Weekly Installment",
    "Total Payments",
    "Total Loan Balance",
    "Loan Officer"
)

Write-Host "Reading input file..."
$lines = Get-Content -Path $inputPath

Write-Host "Writing output with headers..."
$headerLine = '"' + ($headers -join '","') + '"'
$headerLine | Out-File -FilePath $outputPath -Encoding UTF8

# Append all data lines
$lines | Out-File -FilePath $outputPath -Encoding UTF8 -Append

Write-Host "Headers added successfully!"
Write-Host "Output saved to: $outputPath"

# Convert to Excel
Write-Host "Converting to Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$outputExcel = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances_with_headers.xlsx"

try {
    $workbook = $excel.Workbooks.Open($outputPath)
    $workbook.SaveAs($outputExcel, 51)  # 51 = xlOpenXMLWorkbook (xlsx)
    $workbook.Close($false)
    Write-Host "Excel saved to: $outputExcel"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done!"
