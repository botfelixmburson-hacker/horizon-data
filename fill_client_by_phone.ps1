# Fill in missing client IDs by matching phone numbers
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary by phone number
$phoneLookup = @{}
foreach ($client in $completeClients) {
    $phone = $client.'Phone'.Trim()
    $id = $client.'ID Number'.Trim()
    $name = $client.'Client Name'.Trim()

    if ($phone -ne "" -and $phone -ne "0") {
        if (-not $phoneLookup.ContainsKey($phone)) {
            $phoneLookup[$phone] = @{
                ID = $id
                Name = $name
            }
        }
    }
}

Write-Host "Loaded $($phoneLookup.Count) unique phone numbers from complete list"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$foundCount = 0

foreach ($row in $nyahururuData) {
    $clientContact = $row.'CLIENT''S CONTACT'.Trim()
    $clientId = $row.'CLIENTS ID'.Trim()

    # Check if client has phone but missing ID
    if ($clientContact -ne "" -and $clientId -eq "") {
        $updatedCount++

        # Try to find ID by phone number
        if ($phoneLookup.ContainsKey($clientContact)) {
            $clientInfo = $phoneLookup[$clientContact]

            # Fill missing ID
            if ($clientInfo.ID -ne "") {
                $row.'CLIENTS ID' = $clientInfo.ID
                $foundCount++
                Write-Host "  Found ID by phone: $clientContact -> ID: $($clientInfo.ID) (Name: $($clientInfo.Name))"
            }
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Records with phone but missing ID: $updatedCount"
Write-Host "  Found IDs: $foundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
