#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Adds users to the Exempt_OnlineArchive mail-enabled security group from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing user principal names (UPNs) and adds them
    to the specified mail-enabled security group in Exchange Online.
    The CSV file should have a column named 'UserPrincipalName' or 'UPN'.

.PARAMETER CsvPath
    Path to the CSV file containing user principal names.

.PARAMETER GroupIdentity
    The identity of the mail-enabled security group. Default is Exempt_OnlineArchive@starkworkspace.onmicrosoft.com

.PARAMETER WhatIf
    If specified, shows what would happen without making changes.

.EXAMPLE
    .\Add-UsersToExemptGroup.ps1 -CsvPath "C:\users.csv"

.EXAMPLE
    .\Add-UsersToExemptGroup.ps1 -CsvPath "C:\users.csv" -WhatIf

.NOTES
    CSV file format should include a header with either 'UserPrincipalName' or 'UPN'
    Example CSV:
    UserPrincipalName
    user1@domain.com
    user2@domain.com
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    
    [string]$GroupIdentity = "Exempt_OnlineArchive@starkworkspace.onmicrosoft.com",
    
    [switch]$WhatIf
)

try {
    Write-Output "========================================"
    Write-Output "Add Users to Exempt Group Script"
    Write-Output "========================================"
    Write-Output "CSV File: $CsvPath"
    Write-Output "Target Group: $GroupIdentity"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "========================================`n"

    if (-not (Test-Path -Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    Write-Output "Reading CSV file..."
    $users = Import-Csv -Path $CsvPath
    
    if ($users.Count -eq 0) {
        throw "No users found in CSV file"
    }

    $upnColumn = $null
    if ($users[0].PSObject.Properties.Name -contains 'UserPrincipalName') {
        $upnColumn = 'UserPrincipalName'
    }
    elseif ($users[0].PSObject.Properties.Name -contains 'UPN') {
        $upnColumn = 'UPN'
    }
    else {
        throw "CSV file must contain a column named 'UserPrincipalName' or 'UPN'. Found columns: $($users[0].PSObject.Properties.Name -join ', ')"
    }

    Write-Output "Found $($users.Count) user(s) in CSV file"
    Write-Output "Using column: $upnColumn`n"

    Write-Output "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Output "Successfully connected to Exchange Online`n"

    Write-Output "Verifying group exists..."
    try {
        $group = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop
        Write-Output "Group found: $($group.DisplayName) ($($group.PrimarySmtpAddress))"
        Write-Output "Group Type: $($group.RecipientTypeDetails)`n"
    }
    catch {
        throw "Could not find group '$GroupIdentity': $($_.Exception.Message)"
    }

    Write-Output "Retrieving current group members..."
    $currentMembers = @{}
    try {
        $members = Get-DistributionGroupMember -Identity $GroupIdentity -ResultSize Unlimited -ErrorAction Stop
        foreach ($member in $members) {
            if ($member.PrimarySmtpAddress) {
                $currentMembers[$member.PrimarySmtpAddress.ToString().ToLowerInvariant()] = $true
            }
        }
        Write-Output "Current member count: $($currentMembers.Count)`n"
    }
    catch {
        Write-Warning "Could not retrieve current members: $($_.Exception.Message)"
        Write-Output "Continuing anyway...`n"
    }

    Write-Output "Processing users..."
    Write-Output "========================================`n"

    $addedCount = 0
    $alreadyMemberCount = 0
    $errorCount = 0
    $notFoundCount = 0

    foreach ($user in $users) {
        $upn = $user.$upnColumn
        
        if ([string]::IsNullOrWhiteSpace($upn)) {
            Write-Warning "Skipping empty UPN entry"
            $errorCount++
            continue
        }

        $upn = $upn.Trim()

        try {
            $recipient = Get-Recipient -Identity $upn -ErrorAction Stop
            
            if ($currentMembers.ContainsKey($recipient.PrimarySmtpAddress.ToString().ToLowerInvariant())) {
                $alreadyMemberCount++
                Write-Output "Already a member: $($recipient.DisplayName) ($upn)"
                continue
            }

            if ($WhatIf) {
                Write-Output "WhatIf: Would add $($recipient.DisplayName) ($upn) to group"
                Add-DistributionGroupMember -Identity $GroupIdentity -Member $upn -WhatIf -ErrorAction Stop
                $addedCount++
            }
            else {
                Write-Output "Adding: $($recipient.DisplayName) ($upn)"
                Add-DistributionGroupMember -Identity $GroupIdentity -Member $upn -ErrorAction Stop
                $addedCount++
                Write-Output "  Successfully added"
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            if ($errorMessage -match "couldn't be found|not found|does not exist") {
                $notFoundCount++
                Write-Warning "User not found: $upn"
            }
            elseif ($errorMessage -match "already a member") {
                $alreadyMemberCount++
                Write-Output "Already a member: $upn"
            }
            else {
                $errorCount++
                Write-Error "Failed to add $upn : $errorMessage"
            }
        }
    }

    Write-Output "`n========================================"
    Write-Output "Final Summary"
    Write-Output "========================================"
    Write-Output "Total users in CSV: $($users.Count)"
    Write-Output "Successfully added: $addedCount"
    Write-Output "Already members: $alreadyMemberCount"
    Write-Output "Not found: $notFoundCount"
    Write-Output "Errors: $errorCount"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "========================================`n"

    Write-Output "Disconnecting from Exchange Online..."
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Output "Successfully disconnected from Exchange Online"

    Write-Output "`nScript completed successfully"
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Could not disconnect from Exchange Online"
    }
    
    throw
}
