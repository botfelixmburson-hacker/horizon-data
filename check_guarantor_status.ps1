# Check guarantor completion status
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
    $guarantorName = $row.'GUARANTOR''S NAME'.Trim()
    $guarantorContact = $row.'GUARONTOR CONTACT'.Trim()
    $guarantorId = $row.'GUARONTORS ID'.Trim()

    if ($guarantorContact -ne "" -and $guarantorId -ne "") {
        $completeRecords++
    } elseif ($guarantorContact -eq "" -and $guarantorId -eq "") {
        $missingBoth++
        Write-Host "  Missing both: $guarantorName"
    } elseif ($guarantorContact -eq "") {
        $missingContact++
        Write-Host "  Missing contact: $guarantorName (ID: $guarantorId)"
    } elseif ($guarantorId -eq "") {
        $missingId++
        Write-Host "  Missing ID: $guarantorName (Phone: $guarantorContact)"
    }
}

Write-Host "`nSummary:"
Write-Host "  Total records: $totalRecords"
Write-Host "  Complete guarantor info: $completeRecords"
Write-Host "  Missing contact only: $missingContact"
Write-Host "  Missing ID only: $missingId"
Write-Host "  Missing both: $missingBoth"
