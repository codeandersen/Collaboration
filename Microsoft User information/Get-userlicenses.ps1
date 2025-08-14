<# 
REQUIRES:
- Microsoft.Graph module
    Install-Module Microsoft.Graph -Scope CurrentUser

WHAT IT DOES:
- Connects to Microsoft Graph
- Gets all users in M365
- Lists assigned licenses
- Maps SkuPartNumbers to friendly names (custom mapping)
- Outputs CSV report
#>

$ErrorActionPreference = "Stop"

# Mapping SkuPartNumber -> Friendly Name
$SkuPartToName = @{
    "EXCHANGE_S_ENTERPRISE"              = "Exchange Online (Plan 2)"
    "EXCHANGE_S_DESKLESS"                 = "Exchange Online Kiosk"
    "Exchangedeskless"                    = "Exchange Online Kiosk"  # lowercase variant
    "MCOMEETADV"                          = "Microsoft 365 Audio Conferencing"
    "O365_BUSINESS_ESSENTIALS"            = "Microsoft 365 Business Basic"
    "O365_BUSINESS_PREMIUM"               = "Microsoft 365 Business Premium"
    "O365_BUSINESS"                       = "Microsoft 365 Business Standard"
    "M365_E5_NO_TEAMS"                    = "Microsoft 365 E5 EEA (No Teams)"
    "O365_W/O_TEAMS_BUNDLE_M5"            = "Microsoft 365 E5 EEA (No Teams)" # alias
    "MICROSOFTFABRIC_FREE"                = "Microsoft Fabric (Free)"
    "INTUNE_A"                            = "Microsoft Intune Plan 1 Device"
    "POWERAPPS_DEV"                       = "Microsoft Power Apps for Developer"
    "POWER_AUTOMATE_FREE"                 = "Microsoft Power Automate Free"
    "FLOW_FREE"                           = "Microsoft Power Automate Free"  # alias
    "MCOEV"                               = "Microsoft Teams EEA"
    "SPB"                                 = "Microsoft 365 Business Premium" # or Standard, depends on SKU in your tenant
    "Microsoft_Teams_Exploratory_Dept"    = "Microsoft Teams Exploratory"
    "VISIOCLIENT"                         = "Visio Plan 2"
}


Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Directory.Read.All","Organization.Read.All"

# Get all subscribed SKUs for lookup
Write-Host "Fetching subscribed SKUs..." -ForegroundColor Cyan
$subscribedSkus = Get-MgSubscribedSku

# Create a lookup for SkuId -> SkuPartNumber
$SkuIdToSkuPart = @{}
foreach ($sku in $subscribedSkus) {
    $SkuIdToSkuPart[$sku.SkuId.ToString()] = $sku.SkuPartNumber
}

# Get all users
Write-Host "Fetching all users..." -ForegroundColor Cyan
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AssignedLicenses

$total = $users.Count
$counter = 0
$report = @()

foreach ($u in $users) {
    $counter++
    Write-Progress -Activity "Processing users" -Status "$counter of $total : $($u.UserPrincipalName)" -PercentComplete (($counter / $total) * 100)

    $friendlyLicenses = @()

    foreach ($lic in $u.AssignedLicenses) {
        $skuId = $lic.SkuId.ToString()
        if ($SkuIdToSkuPart.ContainsKey($skuId)) {
            $skuPart = $SkuIdToSkuPart[$skuId]
            if ($SkuPartToName.ContainsKey($skuPart)) {
                $friendlyLicenses += $SkuPartToName[$skuPart]
            } else {
                $friendlyLicenses += $skuPart
            }
        } else {
            $friendlyLicenses += "UnknownSKU-$skuId"
        }
    }

    $report += [PSCustomObject]@{
        DisplayName       = $u.DisplayName
        UserPrincipalName = $u.UserPrincipalName
        Licenses          = ($friendlyLicenses | Sort-Object -Unique) -join "; "
    }
}

Write-Progress -Activity "Processing users" -Completed

# Export CSV
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = ".\M365_Users_With_Friendly_Licenses_$ts.csv"
$report | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Host "Done! Report saved to $outFile" -ForegroundColor Green
