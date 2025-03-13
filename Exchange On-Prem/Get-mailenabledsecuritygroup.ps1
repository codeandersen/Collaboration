<#
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
Import-Module ActiveDirectory
#>

#$OU = "OU=Besthuman Groups,OU=Groups,OU=Bestseller Hosting,DC=bestcorp,DC=net"
$besthumangroups = Get-ADObject -SearchBase "OU=Besthuman Groups,OU=Groups,OU=Bestseller Hosting,DC=bestcorp,DC=net" -Properties * -Filter *
#$BHMEG = $besthumangroups | ?{($_.msExchRecipientDisplayType -eq 1073741833) -or ($null -ne $_.mail) -or $_.mail -ne ""}
$BHMEG = $besthumangroups | Where-Object { (($_.msExchRecipientDisplayType -eq 1073741833) -or ($null -ne $_.mail -and "" -ne $_.mail)) }
#$BHMEG1 = $besthumangroups | Where-Object { ($null -ne $_.mail -and "" -ne $_.mail) }

$results = @()
$errorResults = @()
$totalGroups = $BHMEG.Count
$currentGroup = 0

foreach ($group in $BHMEG) {
    $currentGroup++
    Write-Progress -Activity "Processing Groups" -Status "Processing $currentGroup of $totalGroups" -PercentComplete (($currentGroup / $totalGroups) * 100)
    
    try {
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
