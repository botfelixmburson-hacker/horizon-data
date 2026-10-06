# Add Month Disbursed column to nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

# Month name mapping
$monthNames = @{
    1 = "January"
    2 = "February"
    3 = "March"
    4 = "April"
    5 = "May"
    6 = "June"
    7 = "July"
    8 = "August"
    9 = "September"
    10 = "October"
    11 = "November"
    12 = "December"
}

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0

foreach ($row in $nyahururuData) {
    $disbursementDate = $row.'DISBURSEMENT DATE'.Trim()

    if ($disbursementDate -ne "") {
        # Extract month from DD/MM/YYYY format
        if ($disbursementDate -match "/") {
            $parts = $disbursementDate -split "/"
            if ($parts.Count -ge 2) {
                $monthNum = [int]$parts[1]
                if ($monthNum -ge 1 -and $monthNum -le 12) {
                    $monthName = $monthNames[$monthNum]
                    $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value $monthName -Force
                    $updatedCount++
                } else {
                    $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value "" -Force
                }
            } else {
                $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value "" -Force
            }
        } else {
            $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value "" -Force
        }
    } else {
        $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value "" -Force
    }
}

Write-Host "Updated $updatedCount records with month"

# Reorder columns to put Month Disbursed first
$orderedData = $nyahururuData | Select-Object "Month Disbursed", "DISBURSEMENT DATE", "CLIENT'S NAME", "CLIENT'S CONTACT", "CLIENTS ID", "GUARANTOR'S NAME", "GUARONTOR CONTACT", "GUARONTORS ID", "AMOUNT DISBURSED", "REPAYMENT PERIOD", "processing fees", "interest", "OLB", " WEEKLY INSTALMENT", " TOTAL PAYMENTS", "TOTAL LOAN BALANCE", "L.O"

# Export updated data
$orderedData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
