# Remove invalid rows from nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading file..."
$data = Import-Csv -Path $nyahururuPath

$validData = @()
$removedCount = 0

foreach ($row in $data) {
    $month = $row.'Month Disbursed'.Trim()
    $clientName = $row.'CLIENT''S NAME'.Trim()
    $disbursementDate = $row.'DISBURSEMENT DATE'.Trim()
    $clientContact = $row.'CLIENT''S CONTACT'.Trim()
    $clientId = $row.'CLIENTS ID'.Trim()
    $amountDisbursed = $row.'AMOUNT DISBURSED'.Trim()

    # Check if row has valid data
    # Invalid if: no client name OR (no month AND no client contact)
    $hasValidMonth = $month -in @("August", "September", "October")
    $hasClientName = $clientName -ne ""
    $hasClientContact = $clientContact -ne ""
    $hasAmount = $amountDisbursed -ne ""

    # Keep row if it has client name AND (valid month OR client contact OR amount)
    if ($hasClientName -and ($hasValidMonth -or $hasClientContact -or $hasAmount)) {
        $validData += $row
    } else {
        Write-Host "  Removed invalid row: Month='$month', Client='$clientName', Date='$disbursementDate', Contact='$clientContact', Amount='$amountDisbursed'"
        $removedCount++
    }
}

Write-Host "`nRemoved $removedCount invalid rows"
Write-Host "Kept $($validData.Count) valid rows"

# Export cleaned data
$validData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
