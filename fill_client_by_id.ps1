# Fill in missing client phone numbers by matching IDs
$ErrorActionPreference = "Stop"

$nyahururuPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"
$completePath = "C:\Users\admin\Desktop\KashLeo data\complete\kashleo_clients_complete.csv"
$outputPath = "C:\Users\admin\Desktop\KashLeo data\nyahururu\nyahururu_branch_loan_balances.csv"

Write-Host "Loading complete client list..."
$completeClients = Import-Csv -Path $completePath

# Create a lookup dictionary by ID number
$idLookup = @{}
foreach ($client in $completeClients) {
    $id = $client.'ID Number'.Trim()
    $phone = $client.'Phone'.Trim()
    $name = $client.'Client Name'.Trim()

    if ($id -ne "" -and $id -ne "0") {
        if (-not $idLookup.ContainsKey($id)) {
            $idLookup[$id] = @{
                Phone = $phone
                Name = $name
            }
        }
    }
}

Write-Host "Loaded $($idLookup.Count) unique ID numbers from complete list"

Write-Host "Loading Nyahururu loan balances..."
$nyahururuData = Import-Csv -Path $nyahururuPath

$updatedCount = 0
$foundCount = 0

foreach ($row in $nyahururuData) {
    $clientContact = $row.'CLIENT''S CONTACT'.Trim()
    $clientId = $row.'CLIENTS ID'.Trim()

    # Check if client has ID but missing phone
    if ($clientId -ne "" -and $clientContact -eq "") {
        $updatedCount++

        # Try to find phone by ID number
        if ($idLookup.ContainsKey($clientId)) {
            $clientInfo = $idLookup[$clientId]

            # Fill missing phone
            if ($clientInfo.Phone -ne "") {
                $row.'CLIENT''S CONTACT' = $clientInfo.Phone
                $foundCount++
                Write-Host "  Found phone by ID: $clientId -> Phone: $($clientInfo.Phone) (Name: $($clientInfo.Name))"
            }
        }
    }
}

Write-Host "`nSummary:"
Write-Host "  Records with ID but missing phone: $updatedCount"
Write-Host "  Found phones: $foundCount"

# Export updated data
$nyahururuData | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8
Write-Host "`nUpdated file saved to: $outputPath"
