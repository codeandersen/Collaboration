<# 
REQUIRES:
- ExchangeOnlineManagement

WHAT IT DOES:
- Connects to Exchange Online
- Gets all Distribution Groups (excluding Dynamic Distribution Lists)
- Lists all members of each Distribution Group
- Exports the information to a CSV file

OUTPUT:
- .\DistributionGroups_Report_<timestamp>.csv
#>

$ErrorActionPreference = 'Stop'

# Connect to Exchange Online
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all Distribution Groups (excluding Dynamic Distribution Lists)
$distributionGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -ne "DynamicDistributionGroup" }

Write-Host ("Found {0} distribution groups" -f $distributionGroups.Count) -ForegroundColor Cyan

# Initialize report array
$report = @()

foreach ($group in $distributionGroups) {
    $groupType = if ($group.RecipientTypeDetails -eq 'MailUniversalSecurityGroup') { 'Mail-enabled Security Group' } else { 'Distribution Group' }
    Write-Host ("Processing group: {0} ({1})" -f $group.DisplayName, $groupType) -ForegroundColor Yellow

    try {
        # Get members of the distribution group using PrimarySmtpAddress as identifier
        $members = Get-DistributionGroupMember -Identity $group.PrimarySmtpAddress -ResultSize Unlimited

        # Collect member information into a single string
        $memberInfo = ($members | ForEach-Object {
            "$($_.Name) <$($_.PrimarySmtpAddress)>"
        }) -join '; '

        $report += [PSCustomObject]@{
            GroupName       = $group.DisplayName
            GroupEmail      = $group.PrimarySmtpAddress
            MemberInfo      = $memberInfo
            GroupType       = $groupType
        }
    } catch {
        Write-Warning "Failed to get members for group $($group.DisplayName): $($_.Exception.Message)"
    }
}

# Export the report to a CSV file without BOM
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = ".\DistributionGroups_Report_$timestamp.csv"
$report | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

Write-Host "Done. Report saved to: $outputFile" -ForegroundColor Green