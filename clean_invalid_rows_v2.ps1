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

    # Check if row has valid data
    if ($month -in @("August", "September", "October") -and $clientName -ne "") {
        $validData += $row
    } elseif ($month -eq "" -and $clientName -ne "") {
        # Row with empty month but has client name - keep it
        $validData += $row
    } else {
        Write-Host "  Removed invalid row: Month=$month, Client=$clientName, Date=$disbursementDate"
        $removedCount++
    }
}

Write-Host "`nRemoved $removedCount invalid rows"
Write-Host "Kept $($validData.Count) valid rows"

# Export cleaned data
$validData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nCleaned file saved to: $outputPath"
