# Script to update guarantor IDs in CSV files with multiple matching strategies
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
    $guarantorLookup[$entry.Key] = $entry.GuarantorID
}

Write-Host "Loaded $($guarantorLookup.Count) guarantor entries" -ForegroundColor Green

foreach ($targetFile in $targetFiles) {
    Write-Host "`nProcessing: $targetFile" -ForegroundColor Green

    $data = Import-Csv -Path $targetFile
    $updatedCount = 0

    foreach ($record in $data) {
        $clientName = $record.'Client Name'
        $guarantorName = $record.'Guarantor Name'
        $clientContact = $record.'Client Contact'
        $guarantorContact = $record.'Guarantor Contact'
        $currentGuarantorId = $record.'Guarantor ID'

        # Skip if already has guarantor ID
        if ($currentGuarantorId -and $currentGuarantorId -ne "") {
            continue
        }

        # Try multiple matching strategies
        $found = $false

        # Strategy 1: Client Name + Guarantor Name
        $key1 = "$clientName|$guarantorName"
        if ($guarantorLookup.ContainsKey($key1)) {
            $record.'Guarantor ID' = $guarantorLookup[$key1]
            $updatedCount++
            $found = $true
        }

        # Strategy 2: Client Contact + Guarantor Contact
        if (-not $found -and $clientContact -and $guarantorContact) {
            $key2 = "$clientContact|$guarantorContact"
            if ($guarantorLookup.ContainsKey($key2)) {
                $record.'Guarantor ID' = $guarantorLookup[$key2]
                $updatedCount++
                $found = $true
            }
        }

        # Strategy 3: Guarantor Name only
        if (-not $found -and $guarantorName) {
            $key3 = $guarantorName.Trim()
            if ($guarantorLookup.ContainsKey($key3)) {
                $record.'Guarantor ID' = $guarantorLookup[$key3]
                $updatedCount++
                $found = $true
            }
        }
    }

    Write-Host "Updated $updatedCount records" -ForegroundColor Yellow

    # Export updated data
    $data | Export-Csv -Path $targetFile -NoTypeInformation
    Write-Host "Saved: $targetFile" -ForegroundColor Green
}

Write-Host "`nDone!" -ForegroundColor Green
