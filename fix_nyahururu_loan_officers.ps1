# Fix nyahururu_branch_loan_balances.csv with correct data from Excel files
$ErrorActionPreference = "Stop"

$csvFolder = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

$allData = @()

# Standard column names matching the image
$standardHeaders = @(
    "DISBURSEMENT DATE",
    "CLIENT'S NAME",
    "CLIENT'S CONTACT",
    "CLIENTS ID",
    "GUARANTOR'S NAME",
    "GUARONTOR CONTACT",
    "GUARONTORS ID",
    "AMOUNT DISBURSED",
    "REPAYMENT PERIOD",
    "processing fees",
    "interest",
    "OLB",
    " WEEKLY INSTALMENT",
    " TOTAL PAYMENTS",
    "TOTAL LOAN BALANCE",
    "L.O"
)

# Process October
Write-Host "Processing October file..."
$octoberPath = Join-Path $csvFolder "October.csv"
if (Test-Path $octoberPath) {
    $csv = Import-Csv -Path $octoberPath
    $count = 0
    foreach ($row in $csv) {
        $allData += [PSCustomObject]@{
            "DISBURSEMENT DATE" = $row.'DISBURSEMENT DATE'
            "CLIENT'S NAME" = $row.'CLIENT''S NAME'
            "CLIENT'S CONTACT" = $row.'CLIENT''S CONTACT'
            "CLIENTS ID" = $row.'CLIENTS ID'
            "GUARANTOR'S NAME" = $row.'GUARANTOR''S NAME'
            "GUARONTOR CONTACT" = $row.'GUARONTOR CONTACT'
            "GUARONTORS ID" = $row.'GUARANTORS ID'
            "AMOUNT DISBURSED" = $row.'AMOUNT DISBURSED'
            "REPAYMENT PERIOD" = $row.'REPAYMENT PERIOD'
            "processing fees" = $row.'processing fees'
            "interest" = $row.'interest'
            "OLB" = $row.'OLB'
            " WEEKLY INSTALMENT" = $row.' WEEKLY INSTALMENT'
            " TOTAL PAYMENTS" = $row.' TOTAL PAYMENTS'
            "TOTAL LOAN BALANCE" = $row.'TOTAL LOAN BALANCE'
            "L.O" = $row.'L.O'
        }
        $count++
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
        # Use the second name column if first is empty
        $clientName = if ($row.'CLIENT''S NAME' -ne "") { $row.'CLIENT''S NAME' } else { $row.'CLIENTS NAME' }
        $allData += [PSCustomObject]@{
            "DISBURSEMENT DATE" = $row.'DISBURSEMENT DATE'
            "CLIENT'S NAME" = $clientName
            "CLIENT'S CONTACT" = $row.'CLIENTS CONTACT'
            "CLIENTS ID" = $row.'CLIENTS ID'
            "GUARANTOR'S NAME" = $row.'GUARANTOR''S NAME'
            "GUARONTOR CONTACT" = $row.'GUARONTOR CONTACT'
            "GUARONTORS ID" = $row.'GUARONTORS ID'
            "AMOUNT DISBURSED" = $row.'AMOUNT DISBURSED'
            "REPAYMENT PERIOD" = $row.'REPAYMENT PERIOD'
            "processing fees" = $row.'processing fees'
            "interest" = $row.'interest'
            "OLB" = $row.'OLB'
            " WEEKLY INSTALMENT" = $row.' WEEKLY INSTALMENT'
            " TOTAL PAYMENTS" = $row.' TOTAL PAYMENTS'
            "TOTAL LOAN BALANCE" = $row.'TOTAL LOAN BALANCE'
            "L.O" = $row.'L.O'
        }
        $count++
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
        $allData += [PSCustomObject]@{
            "DISBURSEMENT DATE" = $row.'date desbursed'
            "CLIENT'S NAME" = $row.'client name'
            "CLIENT'S CONTACT" = $row.'client number '
            "CLIENTS ID" = $row.'client number'
            "GUARANTOR'S NAME" = $row.'guarantors name'
            "GUARONTOR CONTACT" = $row.'guarantors number'
            "GUARONTORS ID" = $row.' guarantors id '
            "AMOUNT DISBURSED" = $row.'amount disbursed'
            "REPAYMENT PERIOD" = $row.'duration'
            "processing fees" = $row.'processing fee'
            "interest" = $row.'interest'
            "OLB" = $row.'OLB'
            " WEEKLY INSTALMENT" = $row.'weekly installments '
            " TOTAL PAYMENTS" = $row.'amount paid'
            "TOTAL LOAN BALANCE" = $row.'balances'
            "L.O" = $row.'loan officer'
        }
        $count++
    }
    Write-Host "  Added $count records"
}

# Sort by TOTAL LOAN BALANCE (descending)
$allData = $allData | Sort-Object { [double]$_.("TOTAL LOAN BALANCE") } -Descending

Write-Host "`nTotal records: $($allData.Count)"

# Export to CSV
$allData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nOutput saved to: $outputPath"
