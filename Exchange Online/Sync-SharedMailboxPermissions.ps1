<#
.SYNOPSIS
    Automatically manages shared mailbox permissions based on group membership with SM- prefix.

.DESCRIPTION
    This script synchronizes permissions between Microsoft 365 groups and shared mailboxes.
    Groups with prefix "SMBX-xyz" will grant FullAccess (with automapping) to shared mailbox "xyz".
    Users added to the group get permissions, users removed lose permissions.
    
    Uses Managed Identity for authentication (Azure Automation).

.PARAMETER GroupPrefix
    Prefix for groups that manage shared mailbox access. Default: "SMBX-"

.PARAMETER Organization
    Organization domain (e.g., contoso.onmicrosoft.com) - Required for Managed Identity

.PARAMETER WhatIf
    If set to $true, shows what would be done without making any changes (dry-run mode)

.EXAMPLE
    .\Sync-SharedMailboxPermissions.ps1 -GroupPrefix "SMBX-" -Organization "contoso.onmicrosoft.com"

.EXAMPLE
    .\Sync-SharedMailboxPermissions.ps1 -GroupPrefix "SMBX-" -Organization "contoso.onmicrosoft.com" -WhatIf $true

.NOTES

    Requires: ExchangeOnlineManagement module (v3.0+)
    
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$GroupPrefix = "SMBX-",
    
    [Parameter(Mandatory = $false)]
    [string]$Organization = "lemu.onmicrosoft.com",
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"

Write-Output "Checking for ExchangeOnlineManagement module..."
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

if ($module) {
    Write-Output "Found ExchangeOnlineManagement module version $($module.Version)"
    try {
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        Write-Output "Module loaded successfully"
    }
    catch {
        Write-Error "Failed to import ExchangeOnlineManagement module: $_"
        exit 1
    }
}
else {
    Write-Error "ExchangeOnlineManagement module not found. Please install it in the Automation Account."
    exit 1
}

#region Functions

function Connect-ExchangeOnlineAuth {
    <#
    .SYNOPSIS
        Connects to Exchange Online using Managed Identity
    #>
    [CmdletBinding()]
    param()
    
    try {
        Write-Output "Connecting to Exchange Online using Managed Identity..."
        Write-Output "Organization: $Organization"
        Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        Write-Output "Successfully connected using Managed Identity"
    }
    catch {
        Write-Output "ERROR: Failed to connect to Exchange Online: $_"
        Write-Output "ERROR: Exception details: $($_.Exception.Message)"
        throw
    }
}

function Add-SharedMailboxPermission {
    <#
    .SYNOPSIS
        Adds FullAccess and SendAs permissions to a shared mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxIdentity,
        
        [Parameter(Mandatory = $false)]
        [bool]$WhatIfMode = $false
    )
    
    try {
        Write-Verbose "Adding permissions for $UserIdentity to $SharedMailboxIdentity"
        
        if (-not $WhatIfMode) {
            Add-MailboxPermission -Identity $SharedMailboxIdentity `
                -User $UserIdentity `
                -AccessRights FullAccess `
                -InheritanceType All `
                -AutoMapping $true `
                -ErrorAction Stop | Out-Null
            
            Add-RecipientPermission -Identity $SharedMailboxIdentity `
                -Trustee $UserIdentity `
                -AccessRights SendAs `
                -Confirm:$false `
                -ErrorAction Stop | Out-Null
            
            Write-Output "  [+] Successfully added $UserIdentity to $SharedMailboxIdentity (with automapping)"
        }
        else {
            Write-Output "  [WHATIF] Would add $UserIdentity to $SharedMailboxIdentity (with automapping)"
        }
        return $true
    }
    catch {
        Write-Warning "  [!] Failed to add $UserIdentity to $SharedMailboxIdentity - $_"
        return $false
    }
}

function Remove-SharedMailboxPermission {
    <#
    .SYNOPSIS
        Removes FullAccess and SendAs permissions from a shared mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserIdentity,
        
        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxIdentity,
        
        [Parameter(Mandatory = $false)]
        [bool]$WhatIfMode = $false
    )
    
    try {
        Write-Verbose "Removing permissions for $UserIdentity from $SharedMailboxIdentity"
        
        if (-not $WhatIfMode) {
            Remove-MailboxPermission -Identity $SharedMailboxIdentity `
                -User $UserIdentity `
                -AccessRights FullAccess `
                -InheritanceType All `
                -Confirm:$false `
                -ErrorAction Stop | Out-Null
            
            Remove-RecipientPermission -Identity $SharedMailboxIdentity `
                -Trustee $UserIdentity `
                -AccessRights SendAs `
                -Confirm:$false `
                -ErrorAction Stop | Out-Null
            
            Write-Output "  [-] Successfully removed $UserIdentity from $SharedMailboxIdentity"
        }
        else {
            Write-Output "  [WHATIF] Would remove $UserIdentity from $SharedMailboxIdentity"
        }
        return $true
    }
    catch {
        Write-Warning "  [!] Failed to remove $UserIdentity from $SharedMailboxIdentity - $_"
        return $false
    }
}

function Get-SharedMailboxDelegates {
    <#
    .SYNOPSIS
        Gets all users with explicit FullAccess permissions on a shared mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SharedMailboxIdentity
    )
    
    try {
        $permissions = Get-MailboxPermission -Identity $SharedMailboxIdentity -ErrorAction Stop | 
            Where-Object {
                $_.IsInherited -eq $false -and 
                $_.User -ne "NT AUTHORITY\SELF" -and 
                $_.User -notmatch 'S-\d-\d+-\d+-\d+-\d+-\d+-\w+' -and
                $_.AccessRights -contains "FullAccess"
            }
        
        $delegates = @()
        foreach ($perm in $permissions) {
            try {
                $user = Get-User -Identity $perm.User -ErrorAction Stop
                if ($user.RecipientType -eq "UserMailbox") {
                    $delegates += $user.UserPrincipalName
                }
            }
            catch {
                Write-Verbose "Could not resolve user: $($perm.User)"
            }
        }
        
        return $delegates
    }
    catch {
        Write-Warning "Failed to get delegates for $SharedMailboxIdentity - $_"
        return @()
    }
}

function Sync-SharedMailboxGroup {
    <#
    .SYNOPSIS
        Synchronizes permissions between a group and its corresponding shared mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupPrefix,
        
        [Parameter(Mandatory = $false)]
        [bool]$WhatIfMode = $false
    )
    
    Write-Output "DEBUG: Sync-SharedMailboxGroup function started"
    Write-Output "DEBUG: GroupPrefix = $GroupPrefix"
    Write-Output "DEBUG: WhatIfMode = $WhatIfMode"
    
    $stats = @{
        GroupsProcessed = 0
        MailboxesFound = 0
        MailboxesNotFound = 0
        PermissionsAdded = 0
        PermissionsRemoved = 0
        Errors = 0
        Operations = @()
    }
    
    Write-Output "DEBUG: Stats object created"
    
    try {
        Write-Output ""
        Write-Output "Searching for groups with prefix '$GroupPrefix'..."
        $groups = Get-DistributionGroup -Filter "Name -like '$GroupPrefix*'" -ResultSize Unlimited -ErrorAction Stop
        
        if (-not $groups) {
            Write-Warning "No groups found with prefix '$GroupPrefix'"
            return $stats
        }
        
        Write-Output "Found $($groups.Count) group(s) to process"
        Write-Output ""
        
        foreach ($group in $groups) {
            $stats.GroupsProcessed++
            
            # Extract mailbox name from group name (remove prefix)
            $mailboxName = $group.Name -replace "^$([regex]::Escape($GroupPrefix))", ""
            
            # If the extracted name contains @, strip the domain part (e.g., "mailbox@domain.com" -> "mailbox")
            if ($mailboxName -match '@') {
                $mailboxName = $mailboxName -replace '@.*$', ''
            }
            
            Write-Output "[$($stats.GroupsProcessed)/$($groups.Count)] Processing: $($group.Name) -> Mailbox: $mailboxName"
            
            try {
                $sharedMailbox = Get-Mailbox -Identity $mailboxName -ErrorAction Stop | Select-Object -First 1
                
                if ($sharedMailbox.RecipientTypeDetails -ne "SharedMailbox") {
                    Write-Warning "  [!] '$mailboxName' is not a shared mailbox (Type: $($sharedMailbox.RecipientTypeDetails)). Skipping."
                    $stats.Errors++
                    continue
                }
                
                $stats.MailboxesFound++
                $sharedMailboxUPN = [string]$sharedMailbox.UserPrincipalName
                
            }
            catch {
                Write-Warning "  [!] Shared mailbox '$mailboxName' not found. Skipping."
                $stats.MailboxesNotFound++
                continue
            }
            
            try {
                $groupMembers = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop |
                    Where-Object { $_.RecipientType -eq "UserMailbox" } |
                    Select-Object -ExpandProperty PrimarySmtpAddress
                
                Write-Verbose "  Group has $($groupMembers.Count) member(s)"
            }
            catch {
                Write-Warning "  [!] Failed to get members of group $($group.Name) - $_"
                $stats.Errors++
                continue
            }
            
            $currentDelegates = Get-SharedMailboxDelegates -SharedMailboxIdentity $sharedMailboxUPN
            Write-Verbose "  Mailbox has $($currentDelegates.Count) delegate(s)"
            
            if ($groupMembers.Count -eq 0 -and $currentDelegates.Count -gt 0) {
                Write-Output "  [!] Group is empty. Removing all explicit permissions..."
                foreach ($delegate in $currentDelegates) {
                    Write-Output "  [-] Removing $delegate from $sharedMailboxUPN"
                    if (Remove-SharedMailboxPermission -UserIdentity $delegate -SharedMailboxIdentity $sharedMailboxUPN -WhatIfMode $WhatIfMode) {
                        $stats.PermissionsRemoved++
                        $stats.Operations += "REMOVED: $delegate from $mailboxName"
                    }
                    else {
                        $stats.Errors++
                    }
                }
            }
            elseif ($currentDelegates.Count -eq 0 -and $groupMembers.Count -gt 0) {
                Write-Output "  [!] No existing permissions. Adding all group members..."
                foreach ($member in $groupMembers) {
                    Write-Output "  [+] Adding $member to $sharedMailboxUPN"
                    if (Add-SharedMailboxPermission -UserIdentity $member -SharedMailboxIdentity $sharedMailboxUPN -WhatIfMode $WhatIfMode) {
                        $stats.PermissionsAdded++
                        $stats.Operations += "ADDED: $member to $mailboxName"
                    }
                    else {
                        $stats.Errors++
                    }
                }
            }
            elseif ($groupMembers.Count -gt 0 -or $currentDelegates.Count -gt 0) {
                $comparison = Compare-Object -ReferenceObject $currentDelegates -DifferenceObject $groupMembers
                
                if ($comparison) {
                    $usersToAdd = $comparison | Where-Object { $_.SideIndicator -eq "=>" } | Select-Object -ExpandProperty InputObject
                    foreach ($user in $usersToAdd) {
                        Write-Output "  [+] Adding $user to $sharedMailboxUPN"
                        if (Add-SharedMailboxPermission -UserIdentity $user -SharedMailboxIdentity $sharedMailboxUPN -WhatIfMode $WhatIfMode) {
                            $stats.PermissionsAdded++
                            $stats.Operations += "ADDED: $user to $mailboxName"
                        }
                        else {
                            $stats.Errors++
                        }
                    }
                    
                    $usersToRemove = $comparison | Where-Object { $_.SideIndicator -eq "<=" } | Select-Object -ExpandProperty InputObject
                    foreach ($user in $usersToRemove) {
                        Write-Output "  [-] Removing $user from $sharedMailboxUPN"
                        if (Remove-SharedMailboxPermission -UserIdentity $user -SharedMailboxIdentity $sharedMailboxUPN -WhatIfMode $WhatIfMode) {
                            $stats.PermissionsRemoved++
                            $stats.Operations += "REMOVED: $user from $mailboxName"
                        }
                        else {
                            $stats.Errors++
                        }
                    }
                }
                else {
                    Write-Output "  [=] Permissions already in sync"
                }
            }
            
            Write-Output ""
        }
    }
    catch {
        Write-Output "ERROR: Critical error during sync: $_"
        Write-Output "ERROR: Exception Type: $($_.Exception.GetType().FullName)"
        Write-Output "ERROR: Stack Trace: $($_.ScriptStackTrace)"
        $stats.Errors++
    }
    
    return $stats
}

#endregion

#region Main Script

try {
    Write-Output "========================================"
    Write-Output "Shared Mailbox Permission Sync Script"
    if ($WhatIf) {
        Write-Output "*** WHATIF MODE - NO CHANGES WILL BE MADE ***"
    }
    Write-Output "========================================"
    Write-Output "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Group Prefix: $GroupPrefix"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output ""
    
    Connect-ExchangeOnlineAuth
    
    Write-Output "Calling Sync-SharedMailboxGroup function..."
    Write-Output "Parameters: GroupPrefix='$GroupPrefix', WhatIfMode='$WhatIf'"
    
    try {
        $results = Sync-SharedMailboxGroup -GroupPrefix $GroupPrefix -WhatIfMode $WhatIf.IsPresent -ErrorAction Stop
        Write-Output "Sync-SharedMailboxGroup function returned"
        
        if ($null -eq $results) {
            Write-Output "ERROR: Function returned null results"
            $results = @{
                GroupsProcessed = 0
                MailboxesFound = 0
                MailboxesNotFound = 0
                PermissionsAdded = 0
                PermissionsRemoved = 0
                Errors = 1
                Operations = @()
            }
        }
        else {
            Write-Output "Function completed successfully, results object type: $($results.GetType().Name)"
        }
    }
    catch {
        Write-Output "ERROR: Exception calling Sync-SharedMailboxGroup: $_"
        Write-Output "ERROR: Exception Type: $($_.Exception.GetType().FullName)"
        Write-Output "ERROR: Stack Trace: $($_.ScriptStackTrace)"
        throw
    }
    
    Write-Output "========================================"
    Write-Output "Sync Summary"
    Write-Output "========================================"
    Write-Output "Groups Processed:      $($results.GroupsProcessed)"
    Write-Output "Mailboxes Found:       $($results.MailboxesFound)"
    Write-Output "Mailboxes Not Found:   $($results.MailboxesNotFound)"
    Write-Output "Permissions Added:     $($results.PermissionsAdded)"
    Write-Output "Permissions Removed:   $($results.PermissionsRemoved)"
    Write-Output "Errors:                $($results.Errors)"
    
    if ($results.Operations.Count -gt 0) {
        Write-Output ""
        Write-Output "Operations Performed:"
        foreach ($op in $results.Operations) {
            Write-Output "  $op"
        }
    }
    
    Write-Output ""
    Write-Output "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "========================================"
    
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    
    if ($results.Errors -gt 0) {
        exit 1
    }
}
catch {
    Write-Output "FATAL ERROR: Script execution failed: $_"
    Write-Output "FATAL ERROR: Exception Type: $($_.Exception.GetType().FullName)"
    Write-Output "FATAL ERROR: Stack Trace: $($_.ScriptStackTrace)"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

#endregion
