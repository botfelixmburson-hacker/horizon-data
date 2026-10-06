# Combine Nyahururu CSV files and export to Excel
$ErrorActionPreference = "Stop"

$outputCsv = "C:\Users\admin\Desktop\KashLeo data\nyahururu_clients_consolidated.csv"
$outputExcel = "C:\Users\admin\Desktop\KashLeo data\nyahururu_clients_consolidated.xlsx"

$allData = @()

# Process August file
Write-Host "Processing August file..."
$augustPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\aug-balances-nyahururu.csv"
if (Test-Path $augustPath) {
    $csv = Import-Csv -Path $augustPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'balances'
        $balanceNum = if ($balance -ne "" -and $null -ne $balance) { [double]$balance } else { 0 }

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "August"
            "Disbursement Date" = $row.'date desbursed'
            "Client Name" = $row.'client name'
            "Client Contact" = $row.'client number '
            "Client ID" = $row.'client number'
            "Guarantor Name" = $row.'guarantors name'
            "Guarantor Contact" = $row.'guarantors number'
            "Guarantor ID" = $row.' guarantors id '
            "Amount Disbursed" = $row.'amount disbursed'
            "Repayment Period" = $row.'duration'
            "Processing Fees" = $row.'processing fee'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $row.'weekly installments '
            "Total Payments" = $row.'amount paid'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'loan officer'
        }
        $count++
    }
    Write-Host "  Added $count records"
}

# Process October file
Write-Host "Processing October file..."
$octoberPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\october-disbursement-nyahururu.csv"
if (Test-Path $octoberPath) {
    $csv = Import-Csv -Path $octoberPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'TOTAL LOAN BALANCE'
        $balanceNum = if ($balance -ne "" -and $null -ne $balance) { [double]$balance } else { 0 }

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "October"
            "Disbursement Date" = $row.'DISBURSEMENT DATE'
            "Client Name" = $row.'CLIENT''S NAME'
            "Client Contact" = $row.'CLIENT''S CONTACT'
            "Client ID" = $row.'CLIENTS ID'
            "Guarantor Name" = $row.'GUARANTOR''S NAME'
            "Guarantor Contact" = $row.'GUARONTOR CONTACT'
            "Guarantor ID" = $row.'GUARANTORS ID'
            "Amount Disbursed" = $row.'AMOUNT DISBURSED'
            "Repayment Period" = $row.'REPAYMENT PERIOD'
            "Processing Fees" = $row.'processing fees'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $row.' WEEKLY INSTALMENT'
            "Total Payments" = $row.' TOTAL PAYMENTS'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'L.O'
        }
        $count++
    }
    Write-Host "  Added $count records"
}

# Process September file
Write-Host "Processing September file..."
$septemberPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\sep-nyahururu-loan-balances.csv"
if (Test-Path $septemberPath) {
    $csv = Import-Csv -Path $septemberPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'TOTAL LOAN BALANCE'
        $balanceNum = if ($balance -ne "" -and $null -ne $balance) { [double]$balance } else { 0 }

        # Use the second name column if first is empty
        $clientName = if ($row.'CLIENT''S NAME' -ne "" -and $null -ne $row.'CLIENT''S NAME') { $row.'CLIENT''S NAME' } else { $row.'CLIENTS NAME' }

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "September"
            "Disbursement Date" = $row.'DISBURSEMENT DATE'
            "Client Name" = $clientName
            "Client Contact" = $row.'CLIENTS CONTACT'
            "Client ID" = $row.'CLIENTS ID'
            "Guarantor Name" = $row.'GUARANTOR''S NAME'
            "Guarantor Contact" = $row.'GUARONTOR CONTACT'
            "Guarantor ID" = $row.'GUARONTORS ID'
            "Amount Disbursed" = $row.'AMOUNT DISBURSED'
            "Repayment Period" = $row.'REPAYMENT PERIOD'
            "Processing Fees" = $row.'processing fees'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $row.' WEEKLY INSTALMENT'
            "Total Payments" = $row.' TOTAL PAYMENTS'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'L.O'
        }
        $count++
    }
    Write-Host "  Added $count records"
}

# Sort by Total Loan Balance (descending) - 0 balances will be at bottom
Write-Host "`nSorting by Total Loan Balance..."
$allData = $allData | Sort-Object { [double]$_.("Total Loan Balance") } -Descending

Write-Host "Total records: $($allData.Count)"

# Export to CSV first
$allData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
Write-Host "CSV saved to: $outputCsv"

# Convert CSV to Excel
Write-Host "Converting to Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($outputCsv)
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

Write-Host "`nDone!"
