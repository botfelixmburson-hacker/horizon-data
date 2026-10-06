# Fill in missing client ID and phone details from complete client list
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary for clients with normalized names
$clientLookup = @{}
foreach ($client in $completeClients) {
    $name = $client.'Client Name'.Trim()
    $phone = $client.'Phone'.Trim()
    $id = $client.'ID Number'.Trim()

    if ($name -ne "") {
        $normalizedName = ($name -replace '\s+', ' ').ToLower()
        if (-not $clientLookup.ContainsKey($normalizedName)) {
            $clientLookup[$normalizedName] = @{
                OriginalName = $name
                Phone = $phone
                ID = $id
            }
        }
    }
}

Write-Host "Loaded $($clientLookup.Count) unique clients from complete list"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$foundCount = 0

foreach ($row in $nyahururuData) {
    $clientName = $row.'CLIENT''S NAME'.Trim()
    $clientContact = $row.'CLIENT''S CONTACT'.Trim()
    $clientId = $row.'CLIENTS ID'.Trim()

    # Check if client contact or ID is missing
    if ($clientContact -eq "" -or $clientId -eq "") {
        $updatedCount++

        # Normalize the client name for matching
        $normalizedName = ($clientName -replace '\s+', ' ').ToLower()

        # Try exact match first
        if ($clientLookup.ContainsKey($normalizedName)) {
            $clientInfo = $clientLookup[$normalizedName]

            # Fill missing contact
            if ($clientContact -eq "" -and $clientInfo.Phone -ne "") {
                $row.'CLIENT''S CONTACT' = $clientInfo.Phone
            }

            # Fill missing ID
            if ($clientId -eq "" -and $clientInfo.ID -ne "") {
                $row.'CLIENTS ID' = $clientInfo.ID
            }

            $foundCount++
            Write-Host "  Found (exact): $clientName - Phone: $($clientInfo.Phone), ID: $($clientInfo.ID)"
        } else {
            # Try partial match (contains)
            $found = $false
            foreach ($key in $clientLookup.Keys) {
                if ($key -like "*$normalizedName*" -or $normalizedName -like "*$key*") {
                    $clientInfo = $clientLookup[$key]

                    # Fill missing contact
                    if ($clientContact -eq "" -and $clientInfo.Phone -ne "") {
                        $row.'CLIENT''S CONTACT' = $clientInfo.Phone
                    }

                    # Fill missing ID
                    if ($clientId -eq "" -and $clientInfo.ID -ne "") {
                        $row.'CLIENTS ID' = $clientInfo.ID
                    }

                    $foundCount++
                    $found = $true
                    Write-Host "  Found (partial): $clientName matched to $($clientInfo.OriginalName) - Phone: $($clientInfo.Phone), ID: $($clientInfo.ID)"
                    break
                }
            }

            if (-not $found) {
                Write-Host "  Not found: $clientName"
            }
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Records with missing client info: $updatedCount"
Write-Host "  Found in complete list: $foundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
