# Fill in missing guarantor details from complete client list
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary for guarantors
$guarantorLookup = @{}
foreach ($client in $completeClients) {
    $name = $client.'Client Name'.Trim()
    $phone = $client.'Phone'.Trim()
    $id = $client.'ID Number'.Trim()

    if ($name -ne "") {
        if (-not $guarantorLookup.ContainsKey($name)) {
            $guarantorLookup[$name] = @{
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

        # Search in complete client list
        if ($guarantorLookup.ContainsKey($guarantorName)) {
            $guarantorInfo = $guarantorLookup[$guarantorName]

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
            $notFoundCount++
            Write-Host "  Not found: $guarantorName"
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
