# Fill in missing guarantor details with fuzzy name matching
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary for guarantors with normalized names
$guarantorLookup = @{}
foreach ($client in $completeClients) {
    $name = $client.'Client Name'.Trim()
    $phone = $client.'Phone'.Trim()
    $id = $client.'ID Number'.Trim()

    if ($name -ne "") {
        # Normalize name: remove extra spaces, convert to lowercase
        $normalizedName = ($name -replace '\s+', ' ').ToLower()
        if (-not $guarantorLookup.ContainsKey($normalizedName)) {
            $guarantorLookup[$normalizedName] = @{
                OriginalName = $name
                Phone = $phone
                ID = $id
            }
        }
    }
}

Write-Host "Loaded $($guarantorLookup.Count) unique guarantors from complete list"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$foundCount = 0
$notFoundCount = 0

foreach ($row in $nyahururuData) {
    $guarantorName = $row.'GUARANTOR''S NAME'.Trim()
    $guarantorContact = $row.'GUARONTOR CONTACT'.Trim()
    $guarantorId = $row.'GUARONTORS ID'.Trim()

    # Check if guarantor contact or ID is missing
    if ($guarantorContact -eq "" -or $guarantorId -eq "") {
        $updatedCount++

        # Normalize the guarantor name for matching
        $normalizedName = ($guarantorName -replace '\s+', ' ').ToLower()

        # Try exact match first
        if ($guarantorLookup.ContainsKey($normalizedName)) {
            $guarantorInfo = $guarantorLookup[$normalizedName]

            # Fill missing contact
            if ($guarantorContact -eq "" -and $guarantorInfo.Phone -ne "") {
                $row.'GUARONTOR CONTACT' = $guarantorInfo.Phone
            }

            # Fill missing ID
            if ($guarantorId -eq "" -and $guarantorInfo.ID -ne "") {
                $row.'GUARONTORS ID' = $guarantorInfo.ID
            }

            $foundCount++
            Write-Host "  Found (exact): $guarantorName - Phone: $($guarantorInfo.Phone), ID: $($guarantorInfo.ID)"
        } else {
            # Try partial match (contains)
            $found = $false
            foreach ($key in $guarantorLookup.Keys) {
                if ($key -like "*$normalizedName*" -or $normalizedName -like "*$key*") {
                    $guarantorInfo = $guarantorLookup[$key]

                    # Fill missing contact
                    if ($guarantorContact -eq "" -and $guarantorInfo.Phone -ne "") {
                        $row.'GUARONTOR CONTACT' = $guarantorInfo.Phone
                    }

                    # Fill missing ID
                    if ($guarantorId -eq "" -and $guarantorInfo.ID -ne "") {
                        $row.'GUARONTORS ID' = $guarantorInfo.ID
                    }

                    $foundCount++
                    $found = $true
                    Write-Host "  Found (partial): $guarantorName matched to $($guarantorInfo.OriginalName) - Phone: $($guarantorInfo.Phone), ID: $($guarantorInfo.ID)"
                    break
                }
            }

            if (-not $found) {
                $notFoundCount++
                Write-Host "  Not found: $guarantorName"
            }
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Records with missing guarantor info: $updatedCount"
Write-Host "  Found in complete list: $foundCount"
Write-Host "  Not found: $notFoundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
