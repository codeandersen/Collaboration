<# 
REQUIRES:
- ExchangeOnlineManagement

WHAT IT DOES:
- Connects to Exchange Online
- Gets all Microsoft 365 Groups
- Lists all members of each Microsoft 365 Group
- Exports the information to a CSV file

OUTPUT:
- .\M365Groups_Report_<timestamp>.csv
#>

$ErrorActionPreference = 'Stop'

# Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all Microsoft 365 Groups
$m365Groups = Get-UnifiedGroup -ResultSize Unlimited

Write-Host ("Found {0} Microsoft 365 groups" -f $m365Groups.Count) -ForegroundColor Cyan

# Initialize report array
$report = @()

foreach ($group in $m365Groups) {
    Write-Host ("Processing group: {0}" -f $group.DisplayName) -ForegroundColor Yellow

    try {
        # Get members of the Microsoft 365 group
        $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ResultSize Unlimited

        # Collect member information into a single string
        $memberInfo = ($members | ForEach-Object {
            "$($_.Name) <$($_.PrimarySmtpAddress)>"
        }) -join '; '

        $report += [PSCustomObject]@{
            GroupName       = $group.DisplayName
            GroupEmail      = $group.PrimarySmtpAddress
            MemberInfo      = $memberInfo
        }
    } catch {
        Write-Warning "Failed to get members for group $($group.DisplayName): $($_.Exception.Message)"
    }
}

# Export the report to a CSV file without BOM
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = ".\M365Groups_Report_$timestamp.csv"
$report | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Host "Done. Report saved to: $outputFile" -ForegroundColor Green