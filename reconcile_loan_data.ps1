# Reconcile nyahururu_branch_loan_balances.csv with August.csv and September.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$augustPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv\August.csv"
$septemberPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv\September.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading source files..."

# Load August data for lookup
$augustData = @{}
if (Test-Path $augustPath) {
    $csv = Import-Csv -Path $augustPath
    foreach ($row in $csv) {
        $key = "$($row.'client name')|$($row.'date desbursed')"
        $augustData[$key] = @{
            WeeklyInstallment = $row.'weekly installments '
            TotalPayments = $row.'amount paid'
            Balance = $row.'balances'
        }
    }
    Write-Host "  Loaded $($augustData.Count) August records"
}

# Load September data for lookup
$septemberData = @{}
if (Test-Path $septemberPath) {
    $csv = Import-Csv -Path $septemberPath
    foreach ($row in $csv) {
        # Use second name column if first is empty
        $clientName = if ($row.'CLIENT''S NAME' -ne "" -and $null -ne $row.'CLIENT''S NAME') { $row.'CLIENT''S NAME' } else { $row.'CLIENTS NAME' }
        $key = "$($clientName)|$($row.'DISBURSEMENT DATE')"
        $septemberData[$key] = @{
            WeeklyInstallment = $row.' WEEKLY INSTALMENT'
            TotalPayments = $row.' TOTAL PAYMENTS'
            Balance = $row.'TOTAL LOAN BALANCE'
        }
    }
    Write-Host "  Loaded $($septemberData.Count) September records"
}

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$notFoundCount = 0

foreach ($row in $nyahururuData) {
    $clientName = $row.'CLIENT''S NAME'.Trim()
    $disbursementDate = $row.'DISBURSEMENT DATE'.Trim()
    $month = $row.'Month Disbursed'.Trim()

    $key = "$clientName|$disbursementDate"

    # Check if weekly installment or total payments is empty
    $weeklyInstallment = $row.' WEEKLY INSTALMENT'.Trim()
    $totalPayments = $row.' TOTAL PAYMENTS'.Trim()

    if ($weeklyInstallment -eq "" -or $totalPayments -eq "") {
        if ($month -eq "August" -and $augustData.ContainsKey($key)) {
            $sourceData = $augustData[$key]

            if ($weeklyInstallment -eq "" -and $sourceData.WeeklyInstallment -ne "") {
                $row.' WEEKLY INSTALMENT' = $sourceData.WeeklyInstallment
            }

            if ($totalPayments -eq "" -and $sourceData.TotalPayments -ne "") {
                $row.' TOTAL PAYMENTS' = $sourceData.TotalPayments
            }

            # Also update balance if it matches
            if ($sourceData.Balance -ne "") {
                $row.'TOTAL LOAN BALANCE' = $sourceData.Balance
            }

            $updatedCount++
            Write-Host "  Updated August: $clientName - Installment: $($sourceData.WeeklyInstallment), Paid: $($sourceData.TotalPayments)"
        } elseif ($month -eq "September" -and $septemberData.ContainsKey($key)) {
            $sourceData = $septemberData[$key]

            if ($weeklyInstallment -eq "" -and $sourceData.WeeklyInstallment -ne "") {
                $row.' WEEKLY INSTALMENT' = $sourceData.WeeklyInstallment
            }

            if ($totalPayments -eq "" -and $sourceData.TotalPayments -ne "") {
                $row.' TOTAL PAYMENTS' = $sourceData.TotalPayments
            }

            # Also update balance if it matches
            if ($sourceData.Balance -ne "") {
                $row.'TOTAL LOAN BALANCE' = $sourceData.Balance
            }

            $updatedCount++
            Write-Host "  Updated September: $clientName - Installment: $($sourceData.WeeklyInstallment), Paid: $($sourceData.TotalPayments)"
        } else {
            $notFoundCount++
            Write-Host "  Not found: $clientName ($month) - $disbursementDate"
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Updated records: $updatedCount"
Write-Host "  Not found in source: $notFoundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
