# Add Loan Product column and calculate Weekly Installment
$ErrorActionPreference = "Stop"

$inputCsv = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances_with_headers.csv"
$productsCsv = "C:\Users\admin\Desktop\KashLeo data\final data\kashleo_loan_products.csv"
$outputCsv = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances_with_product.csv"
$outputExcel = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances_with_product.xlsx"

# Load loan products
Write-Host "Loading loan products..."
$products = Import-Csv -Path $productsCsv

# Load data
Write-Host "Loading data..."
$data = Import-Csv -Path $inputCsv

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

Write-Host "Processing records..."
$updatedData = @()
$count = 0
$matchedCount = 0

foreach ($row in $data) {
    $amount = $row.'Amount Disbursed'
    $duration = $row.'Repayment Period'

    # Find matching product
    $product = Get-LoanProduct -amount $amount -duration $duration

    # Calculate weekly installment
    $weeklyInstallment = Get-WeeklyInstallment -amount $amount -duration $duration -product $product

    # Create updated row
    $updatedRow = [PSCustomObject]@{
        "Month Disbursed" = $row.'Month Disbursed'
        "Disbursement Date" = $row.'Disbursement Date'
        "Client Name" = $row.'Client Name'
        "Client Contact" = $row.'Client Contact'
        "Client ID" = $row.'Client ID'
        "Guarantor Name" = $row.'Guarantor Name'
        "Guarantor Contact" = $row.'Guarantor Contact'
        "Guarantor ID" = $row.'Guarantor ID'
        "Amount Disbursed" = $row.'Amount Disbursed'
        "Repayment Period" = $row.'Repayment Period'
        "Processing Fees" = $row.'Processing Fees'
        "Interest" = $row.'Interest'
        "OLB" = $row.'OLB'
        "Weekly Installment" = $weeklyInstallment
        "Total Payments" = $row.'Total Payments'
        "Total Loan Balance" = $row.'Total Loan Balance'
        "Loan Officer" = $row.'Loan Officer'
        "Loan Product" = if ($product) { $product.'Product Name' } else { "Unknown" }
    }

    $updatedData += $updatedRow
    $count++

    if ($product) {
        $matchedCount++
    }
}

Write-Host "  Processed $count records"
Write-Host "  Matched $matchedCount records to loan products"

# Export to CSV
Write-Host "Exporting to CSV..."
$updatedData | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

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
