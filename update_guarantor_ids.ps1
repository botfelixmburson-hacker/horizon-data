# Script to update guarantor IDs in CSV files
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv",
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu-arrears-may-jun-jly.csv"
)

Write-Host "Reading guarantor lookup..." -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath

# Create lookup dictionary
$guarantorLookup = @{}
foreach ($entry in $lookupData) {
    $key = "$($entry.ClientName)|$($entry.GuarantorName)"
    $guarantorLookup[$key] = $entry.GuarantorID
}

Write-Host "Loaded $($guarantorLookup.Count) guarantor entries" -ForegroundColor Green

foreach ($targetFile in $targetFiles) {
    Write-Host "`nProcessing: $targetFile" -ForegroundColor Green

    $data = Import-Csv -Path $targetFile
    $updatedCount = 0

    foreach ($record in $data) {
        $clientName = $record.'Client Name'
        $guarantorName = $record.'Guarantor Name'
        $currentGuarantorId = $record.'Guarantor ID'

        # Skip if already has guarantor ID
        if ($currentGuarantorId -and $currentGuarantorId -ne "") {
            continue
        }

        # Look up guarantor ID
        $key = "$clientName|$guarantorName"
        if ($guarantorLookup.ContainsKey($key)) {
            $record.'Guarantor ID' = $guarantorLookup[$key]
            $updatedCount++
        }
    }

    Write-Host "Updated $updatedCount records" -ForegroundColor Yellow

    # Export updated data
    $data | Export-Csv -Path $targetFile -NoTypeInformation
    Write-Host "Saved: $targetFile" -ForegroundColor Green
}

Write-Host "`nDone!" -ForegroundColor Green
