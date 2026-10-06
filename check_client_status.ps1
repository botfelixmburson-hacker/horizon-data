# Check client completion status
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$totalRecords = $nyahururuData.Count
$completeRecords = 0
$missingContact = 0
$missingId = 0
$missingBoth = 0

foreach ($row in $nyahururuData) {
    $clientName = $row.'CLIENT''S NAME'.Trim()
    $clientContact = $row.'CLIENT''S CONTACT'.Trim()
    $clientId = $row.'CLIENTS ID'.Trim()

    if ($clientContact -ne "" -and $clientId -ne "") {
        $completeRecords++
    } elseif ($clientContact -eq "" -and $clientId -eq "") {
        $missingBoth++
        Write-Host "  Missing both: $clientName"
    } elseif ($clientContact -eq "") {
        $missingContact++
        Write-Host "  Missing contact: $clientName (ID: $clientId)"
    } elseif ($clientId -eq "") {
        $missingId++
        Write-Host "  Missing ID: $clientName (Phone: $clientContact)"
    }
}

Write-Host "`nSummary:"
Write-Host "  Total records: $totalRecords"
Write-Host "  Complete client info: $completeRecords"
Write-Host "  Missing contact only: $missingContact"
Write-Host "  Missing ID only: $missingId"
Write-Host "  Missing both: $missingBoth"
