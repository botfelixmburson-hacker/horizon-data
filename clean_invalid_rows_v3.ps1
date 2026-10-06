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

    # Check if row has valid data
    # Valid if: has month AND client name, OR has client name AND contact
    $hasValidMonth = $month -in @("August", "September", "October")
    $hasClientName = $clientName -ne ""
    $hasClientContact = $clientContact -ne ""

    if (($hasValidMonth -and $hasClientName) -or ($hasClientName -and $hasClientContact)) {
        $validData += $row
    } else {
        Write-Host "  Removed invalid row: Month='$month', Client='$clientName', Date='$disbursementDate', Contact='$clientContact'"
        $removedCount++
    }
}

Write-Host "`nRemoved $removedCount invalid rows"
Write-Host "Kept $($validData.Count) valid rows"

# Export cleaned data
$validData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
