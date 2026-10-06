# Add Month Disbursed column to nyahururu_branch_loan_balances.csv
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0

foreach ($row in $nyahururuData) {
    $disbursementDate = $row.'DISBURSEMENT DATE'.Trim()

    if ($disbursementDate -ne "") {
        # Parse the date and extract month
        try {
            # Try DD/MM/YYYY format first
            $dateObj = [DateTime]::ParseExact($disbursementDate, "d/M/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
            $monthName = $dateObj.ToString("MMMM")
            $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value $monthName -Force
            $updatedCount++
        } catch {
            try {
                # Try M/d/yyyy format
                $dateObj = [DateTime]::ParseExact($disbursementDate, "M/d/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                $monthName = $dateObj.ToString("MMMM")
                $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value $monthName -Force
                $updatedCount++
            } catch {
                # If parsing fails, try to extract month from the string
                if ($disbursementDate -match "/") {
                    $parts = $disbursementDate -split "/"
                    if ($parts.Count -ge 2) {
                        $monthNum = [int]$parts[1]
                        if ($monthNum -ge 1 -and $monthNum -le 12) {
                            $monthName = [cultureinfo]::CurrentCulture.DateTimeFormat.GetMonthName($monthNum)
                            $row | Add-Member -MemberType NoteProperty -Name "Month Disbursed" -Value $monthName -Force
                            $updatedCount++
                        }
                    }
                }
            }
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
