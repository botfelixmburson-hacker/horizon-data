# Create arrears file from May, June, July disbursements
$ErrorActionPreference = "Stop"

$productsCsv = "C:\Users\admin\Desktop\KashLeo data\final data\kashleo_loan_products.csv"
$outputCsv = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu-arrears-may-jun-jly.csv"
$outputExcel = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu-arrears-may-jun-jly.xlsx"

# Load loan products
Write-Host "Loading loan products..."
$products = Import-Csv -Path $productsCsv

# Function to find matching loan product
function Get-LoanProduct {
    param($amount, $duration)

    $durationNum = if ($duration -match "(\d+)") { [int]$matches[1] } else { 0 }
    $amountNum = if ($amount -match "^[\d,\.]+$") { [double]($amount -replace ',','') } else { 0 }

    foreach ($product in $products) {
        $minAmount = [double]$product.'Min Loan Amount'
        $maxAmountStr = $product.'Max Loan Amount'
        $maxAmount = if ($maxAmountStr -eq "No Limit") { [double]::MaxValue } else { [double]$maxAmountStr }
        $productDuration = [int]$product.'Duration (Weeks)'

        if ($amountNum -ge $minAmount -and $amountNum -le $maxAmount -and $durationNum -eq $productDuration) {
            return $product
        }
    }

    return $null
}

# Function to calculate weekly installment
function Get-WeeklyInstallment {
    param($amount, $duration, $product)

    $durationNum = if ($duration -match "(\d+)") { [int]$matches[1] } else { 0 }
    $amountNum = if ($amount -match "^[\d,\.]+$") { [double]($amount -replace ',','') } else { 0 }

    if ($product) {
        $interestRate = [double]($product.'Interest Rate (%)' -replace '%','')
        $interestType = $product.'Interest Type'
        $appFeeStr = $product.'Application Fee'
        $appFee = if ($appFeeStr -match '%') { [double]($appFeeStr -replace '%','') } else { [double]$appFeeStr }
        $appFeeType = $product.'Application Fee Type'

        # Calculate interest
        $interest = if ($interestType -eq "Percentage") {
            $amountNum * ($interestRate / 100)
        } else {
            $interestRate
        }

        # Calculate application fee
        $appFeeAmount = if ($appFeeType -eq "Percentage") {
            $amountNum * ($appFee / 100)
        } else {
            $appFee
        }

        # OLB = Amount + Interest
        $olb = $amountNum + $interest

        # Weekly installment = OLB / Duration
        $weeklyInstallment = $olb / $durationNum

        return [math]::Round($weeklyInstallment, 2)
    }

    return ""
}

$allData = @()

# Process May file
Write-Host "Processing May file..."
$mayPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\may-disbursement-nyahururu.csv"
if (Test-Path $mayPath) {
    $csv = Import-Csv -Path $mayPath
    $count = 0
    foreach ($row in $csv) {
        $clientName = $row.'CLIENT''S NAME'
        $balance = $row.'TOTAL LOAN BALANCE'
        $balanceNum = if ($balance -ne "" -and $balance -ne $null) { [double]($balance -replace ',','') } else { 0 }

        # Skip summary rows and records with balance <= 0
        if ($clientName -eq "Total" -or $clientName -eq "" -or $balanceNum -le 0) {
            continue
        }

        $amount = $row.'AMOUNT DISBURSED'
        $duration = $row.'repayment period'
        $product = Get-LoanProduct -amount $amount -duration $duration
        $weeklyInstallment = Get-WeeklyInstallment -amount $amount -duration $duration -product $product

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "May"
            "Disbursement Date" = $row.'Disbursement Date'
            "Client Name" = $row.'CLIENT''S NAME'
            "Client Contact" = $row.'CLIENT''S CONTACT'
            "Client ID" = ""
            "Guarantor Name" = $row.'GUARANTOR''S NAME'
            "Guarantor Contact" = $row.'GUARANTOR''S CONTACT'
            "Guarantor ID" = ""
            "Amount Disbursed" = $amount
            "Repayment Period" = $duration
            "Processing Fees" = $row.'processing fees'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $weeklyInstallment
            "Total Payments" = $row.' TOTAL PAYMENTS'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'L.O'
            "Loan Product" = if ($product) { $product.'Product Name' } else { "Unknown" }
        }
        $count++
    }
    Write-Host "  Added $count records with balance > 0"
}

# Process June file
Write-Host "Processing June file..."
$junePath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\june-disbursement-nyahururu.csv"
if (Test-Path $junePath) {
    $csv = Import-Csv -Path $junePath
    $count = 0
    foreach ($row in $csv) {
        $clientName = $row.'CLIENT''S NAME'
        $balance = $row.'TOTAL LOAN BALANCE'
        $balanceNum = if ($balance -ne "" -and $balance -ne $null) { [double]($balance -replace ',','') } else { 0 }

        # Skip summary rows and records with balance <= 0
        if ($clientName -eq "Total" -or $clientName -eq "" -or $balanceNum -le 0) {
            continue
        }

        $amount = $row.'AMOUNT DISBURSED'
        $duration = $row.'repayment period'
        $product = Get-LoanProduct -amount $amount -duration $duration
        $weeklyInstallment = Get-WeeklyInstallment -amount $amount -duration $duration -product $product

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "June"
            "Disbursement Date" = $row.'Disbursement Date'
            "Client Name" = $row.'CLIENT''S NAME'
            "Client Contact" = $row.'CLIENT''S CONTACT'
            "Client ID" = ""
            "Guarantor Name" = $row.'GUARANTOR''S NAME'
            "Guarantor Contact" = $row.'GUARANTOR''S CONTACT'
            "Guarantor ID" = ""
            "Amount Disbursed" = $amount
            "Repayment Period" = $duration
            "Processing Fees" = $row.'processing fees'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $weeklyInstallment
            "Total Payments" = $row.' TOTAL PAYMENTS'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'L.O'
            "Loan Product" = if ($product) { $product.'Product Name' } else { "Unknown" }
        }
        $count++
    }
    Write-Host "  Added $count records with balance > 0"
}

# Process July file
Write-Host "Processing July file..."
$julyPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\july-disbursement-nyahururu.csv"
if (Test-Path $julyPath) {
    $csv = Import-Csv -Path $julyPath
    $count = 0
    foreach ($row in $csv) {
        $clientName = $row.'CLIENT''S NAME'
        $balance = $row.'TOTAL LOAN BALANCE'
        $balanceNum = if ($balance -ne "" -and $balance -ne $null) { [double]($balance -replace ',','') } else { 0 }

        # Skip summary rows and records with balance <= 0
        if ($clientName -eq "Total" -or $clientName -eq "" -or $balanceNum -le 0) {
            continue
        }

        $amount = $row.'AMOUNT DISBURSED'
        $duration = $row.'repayment period'
        $product = Get-LoanProduct -amount $amount -duration $duration
        $weeklyInstallment = Get-WeeklyInstallment -amount $amount -duration $duration -product $product

        $allData += [PSCustomObject]@{
            "Month Disbursed" = "July"
            "Disbursement Date" = $row.'Disbursement Date'
            "Client Name" = $row.'CLIENT''S NAME'
            "Client Contact" = $row.'CLIENT''S CONTACT'
            "Client ID" = ""
            "Guarantor Name" = $row.'GUARANTOR''S NAME'
            "Guarantor Contact" = $row.'GUARANTOR''S CONTACT'
            "Guarantor ID" = ""
            "Amount Disbursed" = $amount
            "Repayment Period" = $duration
            "Processing Fees" = $row.'processing fees'
            "Interest" = $row.'interest'
            "OLB" = $row.'OLB'
            "Weekly Installment" = $weeklyInstallment
            "Total Payments" = $row.' TOTAL PAYMENTS'
            "Total Loan Balance" = $balanceNum
            "Loan Officer" = $row.'L.O'
            "Loan Product" = if ($product) { $product.'Product Name' } else { "Unknown" }
        }
        $count++
    }
    Write-Host "  Added $count records with balance > 0"
}

Write-Host "Total records with balance > 0: $($allData.Count)"

# Export to CSV
Write-Host "Exporting to CSV..."
$allData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

# Convert to Excel
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

Write-Host "Done!"
