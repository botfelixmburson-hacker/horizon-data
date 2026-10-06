# Combine Nyahururu CSV files with column standardization
$ErrorActionPreference = "Stop"

$csvFolder = "C:\Users\admin\Desktop\KashLeo data\nyahururu_csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\final data\nyahururu_clients_combined.csv"

# Standard column names for output
$standardHeaders = @(
    "Month Disbursed",
    "Disbursement Date",
    "Client Name",
    "Client Contact",
    "Client ID",
    "Guarantor Name",
    "Guarantor Contact",
    "Guarantor ID",
    "Amount Disbursed",
    "Repayment Period",
    "Processing Fees",
    "Interest",
    "OLB",
    "Weekly Installment",
    "Total Payments",
    "Total Loan Balance",
    "Loan Officer"
)

$allData = @()

# Process each file
$files = @(
    @{Month="October"; File="October.csv"; Type="October"},
    @{Month="September"; File="September.csv"; Type="September"},
    @{Month="August"; File="August.csv"; Type="Alternate"}
)

foreach ($fileInfo in $files) {
    $month = $fileInfo.Month
    $filename = $fileInfo.File
    $type = $fileInfo.Type
    $filePath = Join-Path $csvFolder $filename

    Write-Host "Processing $month file: $filename (Type: $type)"

    if (Test-Path $filePath) {
        $lines = Get-Content $filePath
        $headers = $lines[0].Split(",")

        $count = 0
        # Process data rows (skip header)
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $line = $lines[$i].Trim()
            if ($line -eq "") { continue }

            $values = $line.Split(",")

            # Extract data based on file type
            $row = @($month)

            if ($type -eq "October") {
                # Map October format (15 columns)
                # DISBURSEMENT DATE,CLIENT'S NAME,CLIENT'S CONTACT,CLIENTS ID,GUARANTOR'S NAME,GUARONTOR CONTACT,GUARANTORS ID,AMOUNT DISBURSED,REPAYMENT PERIOD,processing fees,interest,OLB, WEEKLY INSTALMENT, TOTAL PAYMENTS,TOTAL LOAN BALANCE,L.O
                $disbDate = if ($values.Count -gt 0) { $values[0] } else { "" }
                $clientName = if ($values.Count -gt 1) { $values[1] } else { "" }
                $clientContact = if ($values.Count -gt 2) { $values[2] } else { "" }
                $clientId = if ($values.Count -gt 3) { $values[3] } else { "" }
                $guarantorName = if ($values.Count -gt 4) { $values[4] } else { "" }
                $guarantorContact = if ($values.Count -gt 5) { $values[5] } else { "" }
                $guarantorId = if ($values.Count -gt 6) { $values[6] } else { "" }
                $amount = if ($values.Count -gt 7) { $values[7] } else { "" }
                $period = if ($values.Count -gt 8) { $values[8] } else { "" }
                $fees = if ($values.Count -gt 9) { $values[9] } else { "" }
                $interest = if ($values.Count -gt 10) { $values[10] } else { "" }
                $olb = if ($values.Count -gt 11) { $values[11] } else { "" }
                $installment = if ($values.Count -gt 12) { $values[12] } else { "" }
                $payments = if ($values.Count -gt 13) { $values[13] } else { "" }
                $balance = if ($values.Count -gt 14) { $values[14] } else { "" }
                $loanOfficer = if ($values.Count -gt 15) { $values[15] } else { "" }
            } elseif ($type -eq "September") {
                # Map September format (16 columns with duplicate name)
                # DISBURSEMENT DATE,CLIENT'S NAME,CLIENTS NAME,CLIENTS CONTACT,CLIENTS ID,GUARANTOR'S NAME,GUARONTOR CONTACT,GUARONTORS ID,AMOUNT DISBURSED,REPAYMENT PERIOD,processing fees,interest,OLB, WEEKLY INSTALMENT, TOTAL PAYMENTS,TOTAL LOAN BALANCE,L.O
                $disbDate = if ($values.Count -gt 0) { $values[0] } else { "" }
                $clientName = if ($values.Count -gt 2) { $values[2] } else { if ($values.Count -gt 1) { $values[1] } else { "" } }
                $clientContact = if ($values.Count -gt 3) { $values[3] } else { "" }
                $clientId = if ($values.Count -gt 4) { $values[4] } else { "" }
                $guarantorName = if ($values.Count -gt 5) { $values[5] } else { "" }
                $guarantorContact = if ($values.Count -gt 6) { $values[6] } else { "" }
                $guarantorId = if ($values.Count -gt 7) { $values[7] } else { "" }
                $amount = if ($values.Count -gt 8) { $values[8] } else { "" }
                $period = if ($values.Count -gt 9) { $values[9] } else { "" }
                $fees = if ($values.Count -gt 10) { $values[10] } else { "" }
                $interest = if ($values.Count -gt 11) { $values[11] } else { "" }
                $olb = if ($values.Count -gt 12) { $values[12] } else { "" }
                $installment = if ($values.Count -gt 13) { $values[13] } else { "" }
                $payments = if ($values.Count -gt 14) { $values[14] } else { "" }
                $balance = if ($values.Count -gt 15) { $values[15] } else { "" }
                $loanOfficer = if ($values.Count -gt 16) { $values[16] } else { "" }
            } else {
                # Map Alternate format (August)
                # date desbursed,client name,client number ,client number,guarantors name,guarantors number, guarantors id ,amount disbursed,duration,processing fee,interest,OLB,weekly installments ,amount paid,balances,loan officer
                $disbDate = if ($values.Count -gt 0) { $values[0] } else { "" }
                $clientName = if ($values.Count -gt 1) { $values[1] } else { "" }
                $clientContact = if ($values.Count -gt 2) { $values[2] } else { "" }
                $clientId = if ($values.Count -gt 3) { $values[3] } else { "" }
                $guarantorName = if ($values.Count -gt 4) { $values[4] } else { "" }
                $guarantorContact = if ($values.Count -gt 5) { $values[5] } else { "" }
                $guarantorId = if ($values.Count -gt 6) { $values[6] } else { "" }
                $amount = if ($values.Count -gt 7) { $values[7] } else { "" }
                $period = if ($values.Count -gt 8) { $values[8] } else { "" }
                $fees = if ($values.Count -gt 9) { $values[9] } else { "" }
                $interest = if ($values.Count -gt 10) { $values[10] } else { "" }
                $olb = if ($values.Count -gt 11) { $values[11] } else { "" }
                $installment = if ($values.Count -gt 12) { $values[12] } else { "" }
                $payments = if ($values.Count -gt 13) { $values[13] } else { "" }
                $balance = if ($values.Count -gt 14) { $values[14] } else { "" }
                $loanOfficer = if ($values.Count -gt 15) { $values[15] } else { "" }
            }

            # Parse balance and filter > 0
            $balanceNum = 0
            if ($balance -ne "") {
                try {
                    $balanceNum = [double]$balance
                } catch {
                    $balanceNum = 0
                }
            }

            if ($balanceNum -gt 0) {
                $row += @($disbDate, $clientName, $clientContact, $clientId, $guarantorName, $guarantorContact, $guarantorId, $amount, $period, $fees, $interest, $olb, $installment, $payments, $balance, $loanOfficer)
                $allData += $row
                $count++
            }
        }

        Write-Host "  Added $count records with balance > 0"
    }
}

# Sort by Total Loan Balance (column index 16, 0-based)
$allData = $allData | Sort-Object { [double]$($_[16]) } -Descending

Write-Host "`nTotal records with balance > 0: $($allData.Count)"

# Write to CSV
$csvContent = $standardHeaders -join ","
foreach ($row in $allData) {
    $csvRow = ($row | ForEach-Object {
        $val = $_.Trim()
        if ($val -match ',') {
            '"{0}"' -f ($val -replace '"', '""')
        } else {
            $val
        }
    }) -join ","
    $csvContent += "`n" + $csvRow
}

$csvContent | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "`nOutput saved to: $outputPath"
