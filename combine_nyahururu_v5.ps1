# Combine Nyahururu CSV files using Import-Csv
$ErrorActionPreference = "Stop"

$csvFolder = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\final data\nyahururu_clients_combined.csv"

$allData = @()

# Process October
Write-Host "Processing October file..."
$octoberPath = Join-Path $csvFolder "October.csv"
if (Test-Path $octoberPath) {
    $csv = Import-Csv -Path $octoberPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'TOTAL LOAN BALANCE'
        if ($balance -ne "" -and [double]$balance -gt 0) {
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
                "Total Loan Balance" = $balance
                "Loan Officer" = $row.'L.O'
            }
            $count++
        }
    }
    Write-Host "  Added $count records"
}

# Process September
Write-Host "Processing September file..."
$septemberPath = Join-Path $csvFolder "September.csv"
if (Test-Path $septemberPath) {
    $csv = Import-Csv -Path $septemberPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'TOTAL LOAN BALANCE'
        if ($balance -ne "" -and [double]$balance -gt 0) {
            # Use the second name column if first is empty
            $clientName = if ($row.'CLIENT''S NAME' -ne "") { $row.'CLIENT''S NAME' } else { $row.'CLIENTS NAME' }
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
                "Total Loan Balance" = $balance
                "Loan Officer" = $row.'L.O'
            }
            $count++
        }
    }
    Write-Host "  Added $count records"
}

# Process August
Write-Host "Processing August file..."
$augustPath = Join-Path $csvFolder "August.csv"
if (Test-Path $augustPath) {
    $csv = Import-Csv -Path $augustPath
    $count = 0
    foreach ($row in $csv) {
        $balance = $row.'balances'
        if ($balance -ne "" -and [double]$balance -gt 0) {
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
                "Total Loan Balance" = $balance
                "Loan Officer" = $row.'loan officer'
            }
            $count++
        }
    }
    Write-Host "  Added $count records"
}

# Sort by Total Loan Balance (descending)
$allData = $allData | Sort-Object { [double]$_.("Total Loan Balance") } -Descending

Write-Host "`nTotal records with balance > 0: $($allData.Count)"

# Export to CSV
$allData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nOutput saved to: $outputPath"
