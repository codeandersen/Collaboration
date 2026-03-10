<#
.SYNOPSIS
    Creates mail-enabled security groups for shared mailboxes that don't have them.

.DESCRIPTION
    This script searches for all shared mailboxes and creates corresponding distribution groups
    with the naming format SMBX-<alias>@lemu.onmicrosoft.com if they don't already exist.
    The groups are hidden from the Global Address List and managed by LEGR@lemu.dk.
    
    Uses Managed Identity for authentication (Azure Automation).

.PARAMETER GroupPrefix
    Prefix for groups that manage shared mailbox access. Default: "SMBX-"

.PARAMETER Organization
    Organization domain (e.g., contoso.onmicrosoft.com) - Required for Managed Identity

.PARAMETER ManagedBy
    Email address of the group manager. Default: "LEGR@lemu.dk"

.PARAMETER WhatIf
    If set, shows what would be done without making any changes (dry-run mode)

.EXAMPLE
    .\Create-SharedMailboxGroups.ps1 -Organization "lemu.onmicrosoft.com"

.EXAMPLE
    .\Create-SharedMailboxGroups.ps1 -Organization "lemu.onmicrosoft.com" -WhatIf

.NOTES
    Requires: ExchangeOnlineManagement module (v3.0+)
    
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$GroupPrefix = "SMBX-",
    
    [Parameter(Mandatory = $false)]
    [string]$Organization = "",
    
    [Parameter(Mandatory = $false)]
    [string]$ManagedBy = "",
    
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

function New-SharedMailboxGroup {
    <#
    .SYNOPSIS
        Creates a distribution group for a shared mailbox
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Alias,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupPrefix,
        
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        
        [Parameter(Mandatory = $true)]
        [string]$ManagedBy,
        
        [Parameter(Mandatory = $false)]
        [switch]$WhatIf
    )
    
    $groupAlias = "$GroupPrefix$Alias"
    $groupName = "$groupAlias@$Organization"
    
    try {
        if ($WhatIf) {
            Write-Output "WHATIF: Would create distribution group: $groupName"
            Write-Output "WHATIF: Alias: $groupAlias"
            Write-Output "WHATIF: ManagedBy: $ManagedBy"
            Write-Output "WHATIF: HiddenFromAddressListsEnabled: True"
            return $true
        }
        
        Write-Output "Creating distribution group: $groupName"
        New-DistributionGroup -Name $groupName -Alias $groupAlias -Type Distribution -ManagedBy $ManagedBy -ErrorAction Stop
        
        Write-Output "Hiding group from Global Address List..."
        Set-DistributionGroup -Identity $groupName -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        
        Write-Output "Successfully created and configured group: $groupName"
        return $true
    }
    catch {
        Write-Output "ERROR: Failed to create group for alias '$Alias': $_"
        Write-Output "ERROR: Exception details: $($_.Exception.Message)"
        return $false
    }
}

#endregion

#region Main Script

try {
    Write-Output "=========================================="
    Write-Output "Starting Shared Mailbox Group Creation"
    Write-Output "=========================================="
    Write-Output "Group Prefix: $GroupPrefix"
    Write-Output "Organization: $Organization"
    Write-Output "Managed By: $ManagedBy"
    Write-Output "WhatIf Mode: $($WhatIf.IsPresent)"
    Write-Output "=========================================="
    
    Connect-ExchangeOnlineAuth
    
    Write-Output "Retrieving all shared mailboxes..."
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    Write-Output "Found $($sharedMailboxes.Count) shared mailbox(es)"
    
    if ($sharedMailboxes.Count -eq 0) {
        Write-Output "No shared mailboxes found. Exiting."
        exit 0
    }
    
    $stats = @{
        Total = $sharedMailboxes.Count
        GroupExists = 0
        GroupCreated = 0
        GroupFailed = 0
    }
    
    foreach ($mailbox in $sharedMailboxes) {
        $alias = $mailbox.Alias
        $groupName = "$GroupPrefix$alias@$Organization"
        
        Write-Output ""
        Write-Output "Processing shared mailbox: $($mailbox.DisplayName) (Alias: $alias)"
        
        try {
            $existingGroup = Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue
            
            if ($existingGroup) {
                Write-Output "Distribution group already exists: $groupName"
                $stats.GroupExists++
            }
            else {
                Write-Output "Distribution group does not exist. Creating..."
                $result = New-SharedMailboxGroup -Alias $alias -GroupPrefix $GroupPrefix -Organization $Organization -ManagedBy $ManagedBy -WhatIf:$WhatIf
                
                if ($result) {
                    $stats.GroupCreated++
                }
                else {
                    $stats.GroupFailed++
                }
            }
        }
        catch {
            Write-Output "ERROR: Failed to process mailbox '$($mailbox.DisplayName)': $_"
            $stats.GroupFailed++
        }
    }
    
    Write-Output ""
    Write-Output "=========================================="
    Write-Output "Summary"
    Write-Output "=========================================="
    Write-Output "Total shared mailboxes: $($stats.Total)"
    Write-Output "Groups already existed: $($stats.GroupExists)"
    Write-Output "Groups created: $($stats.GroupCreated)"
    Write-Output "Groups failed: $($stats.GroupFailed)"
    Write-Output "=========================================="
    
    if ($stats.GroupFailed -gt 0) {
        Write-Output "WARNING: Some groups failed to create. Review the log above for details."
    }
    
    Write-Output "Script completed successfully"
}
catch {
    Write-Output "FATAL ERROR: Script execution failed: $_"
    Write-Output "ERROR: Exception details: $($_.Exception.Message)"
    exit 1
}
finally {
    try {
        Write-Output "Disconnecting from Exchange Online..."
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Output "Disconnected successfully"
    }
    catch {
        Write-Output "WARNING: Failed to disconnect cleanly: $_"
    }
}

#endregion
