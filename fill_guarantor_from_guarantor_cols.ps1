# Fill in missing guarantor details from the guarantor columns in complete client list
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary for guarantors from the guarantor columns
$guarantorLookup = @{}
foreach ($client in $completeClients) {
    $guarantorName = $client.'Guarantor Name'.Trim()
    $guarantorContact = $client.'Guarantor Contact'.Trim()
    $guarantorId = $client.'Guarantor ID'.Trim()

    if ($guarantorName -ne "") {
        $normalizedName = ($guarantorName -replace '\s+', ' ').ToLower()
        if (-not $guarantorLookup.ContainsKey($normalizedName)) {
            $guarantorLookup[$normalizedName] = @{
                Phone = $guarantorContact
                ID = $guarantorId
            }
        }
    }
}

Write-Host "Loaded $($guarantorLookup.Count) unique guarantors from guarantor columns in complete list"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$foundCount = 0

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
            Write-Host "  Found: $guarantorName - Phone: $($guarantorInfo.Phone), ID: $($guarantorInfo.ID)"
        } else {
            # Try partial match
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
                    Write-Host "  Found (partial): $guarantorName matched to $key - Phone: $($guarantorInfo.Phone), ID: $($guarantorInfo.ID)"
                    break
                }
            }
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Records with missing guarantor info: $updatedCount"
Write-Host "  Found in complete list: $foundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
