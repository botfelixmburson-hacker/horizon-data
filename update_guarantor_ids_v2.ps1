# Script to update guarantor IDs in CSV files by guarantor name only
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv",
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu-arrears-may-jun-jly.csv"
)

Write-Host "Reading guarantor lookup..." -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath

# Create lookup dictionary by guarantor name only
$guarantorLookup = @{}
foreach ($entry in $lookupData) {
    $guarantorName = $entry.GuarantorName.Trim()
    if (-not $guarantorLookup.ContainsKey($guarantorName)) {
        $guarantorLookup[$guarantorName] = $entry.GuarantorID
    }
}

Write-Host "Loaded $($guarantorLookup.Count) unique guarantor entries" -ForegroundColor Green

foreach ($targetFile in $targetFiles) {
    Write-Host "`nProcessing: $targetFile" -ForegroundColor Green

    $data = Import-Csv -Path $targetFile
    $updatedCount = 0

    foreach ($record in $data) {
        $guarantorName = $record.'Guarantor Name'
        $currentGuarantorId = $record.'Guarantor ID'

        # Skip if already has guarantor ID
        if ($currentGuarantorId -and $currentGuarantorId -ne "") {
            continue
        }

        # Skip if no guarantor name
        if (-not $guarantorName -or $guarantorName -eq "") {
            continue
        }

        # Look up guarantor ID by name (trim and normalize)
        $guarantorNameKey = $guarantorName.Trim()
        if ($guarantorLookup.ContainsKey($guarantorNameKey)) {
            $record.'Guarantor ID' = $guarantorLookup[$guarantorNameKey]
            $updatedCount++
        }
    }

    Write-Host "Updated $updatedCount records" -ForegroundColor Yellow

    # Export updated data
    $data | Export-Csv -Path $targetFile -NoTypeInformation
    Write-Host "Saved: $targetFile" -ForegroundColor Green
}

Write-Host "`nDone!" -ForegroundColor Green
