<#
# Uncomment these lines if not already connected to Exchange
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
Import-Module ActiveDirectory
#>

# Define target OU
$targetOU = "OU=xxxx Groups,OU=Groups,OU=xxxxx Hosting,DC=xxxxx,DC=net"

# Get all mail-enabled security groups from the specified OU
Write-Host "Getting all mail-enabled security groups from $targetOU" -ForegroundColor Cyan
$mailEnabledSecurityGroups = Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -OrganizationalUnit $targetOU

# If Get-Recipient with OrganizationalUnit parameter doesn't work in your environment, use this alternative approach:
# $allGroups = Get-ADGroup -SearchBase $targetOU -Filter * -Properties mail, groupType
# $mailEnabledSecurityGroups = $allGroups | Where-Object { 
#    ($null -ne $_.mail -and $_.mail -ne "") -and 
#    ($_.groupType -band 0x80000000) -and 
#    ($_.groupType -band 0x00000008)
# }

$results = @()
$errorResults = @()
$totalGroups = $mailEnabledSecurityGroups.Count
$currentGroup = 0

Write-Host "Found $totalGroups mail-enabled security groups to process" -ForegroundColor Yellow

foreach ($group in $mailEnabledSecurityGroups) {
    $currentGroup++
    Write-Progress -Activity "Processing Mail-Enabled Security Groups" -Status "Processing $currentGroup of $totalGroups" -PercentComplete (($currentGroup / $totalGroups) * 100)
    
    try {
        # Get the group members
        $members = Get-ADGroupMember -Identity $group.DistinguishedName
        $hasMailEnabledMember = $false
        
        foreach ($member in $members) {
            if ($member.objectClass -eq "user" -or $member.objectClass -eq "inetOrgPerson") {
                $hasMailEnabledMember = $true
                break
            } elseif ($member.objectClass -eq "group") {
                $memberDetails = Get-ADGroup -Identity $member.DistinguishedName -Properties mail
                if ($null -ne $memberDetails.mail -and "" -ne $memberDetails.mail) {
                    $hasMailEnabledMember = $true
                    break
                }
            }
        }
        
        if ($hasMailEnabledMember -eq $false) {
            Write-Host $group.DistinguishedName -ForegroundColor Green
            $results += $group
        }
    }
    catch {
        Write-Host "Error processing group: $($group.DistinguishedName)" -ForegroundColor Red
        Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
        $errorResults += [PSCustomObject]@{
            DistinguishedName = $group.DistinguishedName
            ErrorMessage = $_.Exception.Message
            TimeStamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$successFilename = "MailEnabledGroupsWithNonMailEnabledMembers_$timestamp.csv"
$errorFilename = "MailEnabledGroupsWithNonMailEnabledMembersGroupNotFound_$timestamp.csv"

$results | Export-Csv -Path $successFilename -NoTypeInformation -NoClobber -Delimiter '¤' -Encoding Unicode
if ($errorResults.Count -gt 0) {
    $errorResults | Export-Csv -Path $errorFilename -NoTypeInformation -NoClobber -Delimiter '¤' -Encoding Unicode
}

Write-Output "Script completed. Check the CSV files:"
Write-Output "- Success results: $successFilename"
if ($errorResults.Count -gt 0) {
    Write-Output "- Error results: $errorFilename"
}
