# Script to update guarantor IDs in CSV files with fuzzy matching
$lookupPath = "C:\Users\admin\Desktop\KashLeo data\guarantor_lookup.csv"
$targetFiles = @(
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu_branch_loan_balances.csv",
    "C:\Users\admin\Desktop\KashLeo data\complete\nyahururu-arrears-may-jun-jly.csv"
)

Write-Host "Reading guarantor lookup..." -ForegroundColor Green
$lookupData = Import-Csv -Path $lookupPath

# Create lookup dictionary with normalized keys (lowercase, trimmed)
$guarantorLookup = @{}
foreach ($entry in $lookupData) {
    $normalizedName = $entry.GuarantorName.Trim().ToLower()
    if (-not $guarantorLookup.ContainsKey($normalizedName)) {
        $guarantorLookup[$normalizedName] = $entry.GuarantorID
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

        # Normalize guarantor name for matching
        $normalizedName = $guarantorName.Trim().ToLower()

        # Look up guarantor ID
        if ($guarantorLookup.ContainsKey($normalizedName)) {
            $record.'Guarantor ID' = $guarantorLookup[$normalizedName]
            $updatedCount++
        }
    }

    Write-Host "Updated $updatedCount records" -ForegroundColor Yellow

    # Export updated data
    $data | Export-Csv -Path $targetFile -NoTypeInformation
    Write-Host "Saved: $targetFile" -ForegroundColor Green
}

Write-Host "`nDone!" -ForegroundColor Green
