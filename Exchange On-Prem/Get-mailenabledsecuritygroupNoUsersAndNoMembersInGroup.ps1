<#
# Uncomment these lines if not already connected to Exchange
if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
Import-Module ActiveDirectory
#>

# Define target OU
$targetOU = "OU=xxxx Groups,OU=Groups,OU=xxxx Hosting,DC=xxxxx,DC=net"

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

Write-Host "Found $($mailEnabledSecurityGroups.Count) mail-enabled security groups to process" -ForegroundColor Yellow

# Function to recursively check group membership
function Test-GroupHasEmptyNestedGroups {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupDN,
        
        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.HashSet[string]]$ProcessedGroups = $null
    )
    
    # Initialize the hash set if not provided (first call)
    if ($null -eq $ProcessedGroups) {
        $ProcessedGroups = New-Object System.Collections.Generic.HashSet[string]
    }
    
    # Skip if we've already processed this group (prevents infinite recursion)
    if (-not $ProcessedGroups.Add($GroupDN)) {
        return $false
    }

    try {
        # Get direct members of this group
        $members = Get-ADGroupMember -Identity $GroupDN -ErrorAction Stop
        
        # If the group has no members at all, it's empty
        if ($members.Count -eq 0) {
            return $true
        }
        
        # Check if there are any non-group members (users, contacts, inetOrgPerson, etc.)
        $nonGroupMembers = $members | Where-Object { $_.objectClass -ne "group" }
        if ($nonGroupMembers.Count -gt 0) {
            return $false
        }
        
        # Now check all the nested groups recursively
        foreach ($nestedGroup in $members) {
            if (Test-GroupHasEmptyNestedGroups -GroupDN $nestedGroup.DistinguishedName -ProcessedGroups $ProcessedGroups) {
                return $true
            }
        }
    }
    catch {
        Write-Host "Error checking group $GroupDN`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    # If we get here, no empty nested groups were found
    return $false
}

# Function to check if a group only has other groups as members (no direct users or inetOrgPerson)
function Test-GroupOnlyHasGroupMembers {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupDN
    )
    
    try {
        # Get the members of the group
        $members = Get-ADGroupMember -Identity $GroupDN -ErrorAction Stop
        
        if ($members.Count -eq 0) {
            # No members at all
            return $false
        }
        
        # Explicitly check for inetOrgPerson objects since they need to be excluded
        $hasInetOrgPerson = $members | Where-Object { $_.objectClass -eq "inetOrgPerson" } | Select-Object -First 1
        
        # If we found any inetOrgPerson objects, we should exclude this group
        if ($null -ne $hasInetOrgPerson) {
            return $false
        }
        
        # Check if there are any non-group members
        $nonGroupMembers = $members | Where-Object { $_.objectClass -ne "group" }
        return ($nonGroupMembers.Count -eq 0)
    }
    catch {
        Write-Host "Error checking group members for $GroupDN`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Results collections
$results = @()
$errorResults = @()
$totalGroups = $mailEnabledSecurityGroups.Count
$currentGroup = 0

# Process each mail-enabled security group
foreach ($group in $mailEnabledSecurityGroups) {
    $currentGroup++
    Write-Progress -Activity "Processing Mail-Enabled Security Groups" -Status "Processing $currentGroup of $totalGroups" -PercentComplete (($currentGroup / $totalGroups) * 100)
    
    try {
        # Check if the group only has groups as members
        $onlyHasGroupMembers = Test-GroupOnlyHasGroupMembers -GroupDN $group.DistinguishedName
        
        if ($onlyHasGroupMembers) {
            # Check if any of the nested groups are empty
            $hasEmptyNestedGroups = Test-GroupHasEmptyNestedGroups -GroupDN $group.DistinguishedName
            
            if ($hasEmptyNestedGroups) {
                Write-Host "Found mail-enabled security group with empty nested groups: $($group.Name)" -ForegroundColor Green
                
                # Get additional properties for reporting
                $adGroup = Get-ADGroup -Identity $group.DistinguishedName -Properties mail, displayName, description, managedBy
                
                $results += [PSCustomObject]@{
                    Name = $group.Name
                    DisplayName = $adGroup.displayName
                    Email = $adGroup.mail
                    Description = $adGroup.description
                    DistinguishedName = $group.DistinguishedName
                    ManagedBy = $adGroup.managedBy
                    TimeFound = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                }
            }
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
$successFilename = "MailEnabledSecurityGroups_WithEmptyNestedGroups_$timestamp.csv"
$errorFilename = "MailEnabledSecurityGroups_Errors_$timestamp.csv"

if ($results.Count -gt 0) {
    Write-Host "Found $($results.Count) mail-enabled security groups with empty nested groups" -ForegroundColor Yellow
    $results | Export-Csv -Path $successFilename -NoTypeInformation -NoClobber -Delimiter '¤' -Encoding Unicode
}
else {
    Write-Host "No mail-enabled security groups with empty nested groups were found" -ForegroundColor Yellow
}

if ($errorResults.Count -gt 0) {
    $errorResults | Export-Csv -Path $errorFilename -NoTypeInformation -NoClobber -Delimiter '¤' -Encoding Unicode
}

Write-Output "Script completed. Check the CSV files:"
if ($results.Count -gt 0) {
    Write-Output "- Results: $successFilename"
}
if ($errorResults.Count -gt 0) {
    Write-Output "- Error results: $errorFilename"
}
