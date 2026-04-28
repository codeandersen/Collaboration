<#
.SYNOPSIS
    Syncs mailbox users into mail-enabled security groups based on their Exchange Office attribute.

.DESCRIPTION
    This Azure Automation runbook performs a full sync (add and remove) of mailbox users into
    mail-enabled security groups based on each user's Office attribute.

    Groups follow the naming pattern: {Prefix}_{OfficeValue}_{Suffix}@{Domain}
    For example, a user with Office = "D137" and default settings will be synced into:
        ACL_D137_BEST@contoso.dk
        ACL_D137_INFO@contoso.dk
        ACL_D137_SALG@contoso.dk

    The script:
    - Adds users to groups they should belong to (based on Office value)
    - Removes users from groups they should no longer belong to
    - Skips users with an empty Office attribute (logged)
    - Skips groups that are synced from on-premises / DirSynced (logged)
    - Supports WhatIf (dry-run) and DebugMode (verbose logging)

    Designed to run in Azure Automation with Managed Identity authentication.

.PARAMETER ListPrefix
    Prefix used in the group naming pattern. Default: "ACL"

.PARAMETER Suffixes
    Array of suffixes to sync. Each suffix corresponds to one group per Office value.
    Default: @("BEST","INFO","SALG")

.PARAMETER Domain
    Email domain for the group primary SMTP address. Default: "contoso.dk"

.EXAMPLE
    .\Sync-OfficeSecurityGroups.ps1
    Runs a full sync with default settings. Configure Organization, WhatIfMode, and DebugMode
    in the #region Configuration section at the top of the script.

.EXAMPLE
    .\Sync-OfficeSecurityGroups.ps1 -ListPrefix "DL" -Suffixes @("ALL","MGMT") -Domain "contoso.com"
    Custom prefix, suffixes, and domain.

.NOTES
    Requires: ExchangeOnlineManagement module (v3.0+) installed in the Automation Account
    Authentication: Managed Identity (system-assigned)
    Role required: Exchange Administrator assigned to the Managed Identity

    Reference: https://learn.microsoft.com/en-us/powershell/exchange/connect-exo-powershell-managed-identity
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$ListPrefix = "ACL",

    [Parameter(Mandatory = $false)]
    [string[]]$Suffixes = @("BEST", "INFO", "SALG"),

    [Parameter(Mandatory = $false)]
    [string]$Domain = "contoso.dk"
)

$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"

#region Configuration
$Organization = "contoso.onmicrosoft.com"    # Tenant domain for Managed Identity authentication
$WhatIfMode   = $false                       # Set to $true for dry-run (no changes made)
$DebugMode    = $false                       # Set to $true for verbose debug logging
#endregion

#region Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped message to the Azure Automation output stream.
    #>
    param (
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Debug')]
        [string]$Level = 'Info'
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    switch ($Level) {
        'Warning' { Write-Warning "[$ts] $Message" }
        'Error'   { Write-Output "[$ts] [ERROR] $Message" }
        'Debug'   {
            if ($script:DebugMode) {
                Write-Output "[$ts] [DEBUG] $Message"
            }
        }
        default   { Write-Output "[$ts] [INFO] $Message" }
    }
}

function Connect-ExchangeOnlineAuth {
    <#
    .SYNOPSIS
        Connects to Exchange Online using Managed Identity.
    #>
    [CmdletBinding()]
    param()

    try {
        Write-Log "Connecting to Exchange Online using Managed Identity..."
        Write-Log "Organization: $Organization"
        Connect-ExchangeOnline -ManagedIdentity -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online"
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-AllMailboxUsers {
    <#
    .SYNOPSIS
        Fetches all mailbox users with their Office and PrimarySmtpAddress in a single call.
    #>
    [CmdletBinding()]
    param()

    Write-Log "Fetching all mailbox users (this may take a moment)..."

    try {
        $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties Office, PrimarySmtpAddress, DisplayName -ErrorAction Stop)
        Write-Log "Retrieved $($mailboxes.Count) mailbox user(s)"
        return $mailboxes
    }
    catch {
        Write-Log "Failed to retrieve mailbox users: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Get-AllMatchingGroups {
    <#
    .SYNOPSIS
        Fetches all mail-enabled security groups matching the prefix pattern in a single call.
    #>
    [CmdletBinding()]
    param(
        [string]$Prefix
    )

    Write-Log "Fetching distribution groups matching prefix '$Prefix'..."

    try {
        $groups = @(Get-DistributionGroup -Filter "Name -like '$Prefix`_*'" -ResultSize Unlimited -ErrorAction Stop)
        Write-Log "Retrieved $($groups.Count) matching group(s)"
        return $groups
    }
    catch {
        Write-Log "Failed to retrieve distribution groups: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Parse-GroupName {
    <#
    .SYNOPSIS
        Parses a group name to extract the Office value and Suffix.
        Expected format: {Prefix}_{OfficeValue}_{Suffix}
    #>
    [CmdletBinding()]
    param(
        [string]$GroupName,
        [string]$Prefix,
        [string[]]$ValidSuffixes
    )

    # Remove the prefix and leading underscore
    $remainder = $GroupName
    if ($remainder.StartsWith("${Prefix}_")) {
        $remainder = $remainder.Substring($Prefix.Length + 1)
    }
    else {
        return $null
    }

    # Split into segments and search right-to-left for a valid suffix.
    # This handles group names with trailing hashes (e.g., ACL_D859_BEST_bc22ab89e9)
    $segments = $remainder -split '_'

    for ($i = $segments.Count - 1; $i -ge 1; $i--) {
        if ($ValidSuffixes -contains $segments[$i]) {
            $officeValue = ($segments[0..($i - 1)]) -join '_'
            $suffix = $segments[$i]
            return @{
                OfficeValue = $officeValue
                Suffix      = $suffix
            }
        }
    }

    return $null
}

#endregion

#region Main Script

try {
    Write-Log "========================================"
    Write-Log "Sync Office Security Groups"
    if ($WhatIfMode) {
        Write-Log "*** WHATIF MODE - NO CHANGES WILL BE MADE ***"
    }
    if ($DebugMode) {
        Write-Log "*** DEBUG MODE ENABLED ***"
    }
    Write-Log "========================================"
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Log "Organization: $Organization"
    Write-Log "ListPrefix: $ListPrefix"
    Write-Log "Suffixes: $($Suffixes -join ', ')"
    Write-Log "Domain: $Domain"
    Write-Log "WhatIf: $WhatIfMode"
    Write-Log "DebugMode: $DebugMode"
    Write-Log "========================================"
    Write-Log ""

    # Check for ExchangeOnlineManagement module
    Write-Log "Checking for ExchangeOnlineManagement module..."
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

    if ($module) {
        Write-Log "Found ExchangeOnlineManagement module version $($module.Version)"
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        Write-Log "Module loaded successfully"
    }
    else {
        Write-Log "ExchangeOnlineManagement module not found. Please install it in the Automation Account." -Level Error
        exit 1
    }

    # Connect to Exchange Online
    Connect-ExchangeOnlineAuth

    # Statistics
    $stats = @{
        TotalUsers         = 0
        UsersSkippedNoOffice = 0
        UniqueOfficeValues  = 0
        GroupsProcessed    = 0
        GroupsDirSynced    = 0
        GroupsSkippedParse = 0
        MembersAdded       = 0
        MembersRemoved     = 0
        MembersAlreadySync = 0
        Errors             = 0
    }

    # -------------------------------------------------------
    # Step 1: Fetch all mailbox users and build Office lookup
    # -------------------------------------------------------
    $allMailboxes = Get-AllMailboxUsers
    $stats.TotalUsers = $allMailboxes.Count

    # Build hashtable: OfficeValue -> [list of PrimarySmtpAddress]
    $officeLookup = @{}
    $skippedNoOfficeCount = 0

    foreach ($mbx in $allMailboxes) {
        $officeValue = $mbx.Office
        $smtp = $mbx.PrimarySmtpAddress

        if ([string]::IsNullOrWhiteSpace($officeValue)) {
            $skippedNoOfficeCount++
            Write-Log "Skipping user '$($mbx.DisplayName)' ($smtp) - Office attribute is empty" -Level Debug
            continue
        }

        # Split comma-separated Office values (e.g., "D137, D187, D297")
        $officeKeys = $officeValue -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
        $smtpLower = $smtp.ToString().ToLowerInvariant()

        foreach ($officeKey in $officeKeys) {
            if (-not $officeLookup.ContainsKey($officeKey)) {
                $officeLookup[$officeKey] = [System.Collections.Generic.List[string]]::new()
            }
            $officeLookup[$officeKey].Add($smtpLower)
        }

        Write-Log "User '$($mbx.DisplayName)' ($smtp) -> Office: $($officeKeys -join ', ')" -Level Debug
    }

    $stats.UsersSkippedNoOffice = $skippedNoOfficeCount
    $stats.UniqueOfficeValues = $officeLookup.Count

    Write-Log ""
    Write-Log "Office lookup built: $($officeLookup.Count) unique Office value(s), $skippedNoOfficeCount user(s) skipped (empty Office)"

    # -------------------------------------------------------
    # Step 2: Fetch all matching groups
    # -------------------------------------------------------
    $allGroups = Get-AllMatchingGroups -Prefix $ListPrefix

    if ($allGroups.Count -eq 0) {
        Write-Log "No groups found matching prefix '$ListPrefix'. Nothing to sync." -Level Warning
        Write-Log "Expected group naming pattern: ${ListPrefix}_<OfficeValue>_<Suffix>" -Level Warning
    }

    # -------------------------------------------------------
    # Step 3: Process each group
    # -------------------------------------------------------
    $groupIndex = 0

    foreach ($group in $allGroups) {
        $groupIndex++
        $groupName = $group.Name
        $stats.GroupsProcessed++

        Write-Log ""
        Write-Log "[$groupIndex/$($allGroups.Count)] Processing group: $groupName ($($group.PrimarySmtpAddress))"

        # Parse group name to extract OfficeValue and Suffix
        $parsed = Parse-GroupName -GroupName $groupName -Prefix $ListPrefix -ValidSuffixes $Suffixes

        if ($null -eq $parsed) {
            Write-Log "  Could not parse group name '$groupName' into expected pattern ${ListPrefix}_<OfficeValue>_<Suffix>. Skipping." -Level Warning
            $stats.GroupsSkippedParse++
            Write-Log "  Group name did not match pattern. Prefix='$ListPrefix', ValidSuffixes='$($Suffixes -join ',')', GroupName='$groupName'" -Level Debug
            continue
        }

        $officeValue = $parsed.OfficeValue
        $suffix = $parsed.Suffix
        Write-Log "  Parsed -> Office: $officeValue, Suffix: $suffix" -Level Debug

        # Verify cloud-only (not DirSynced) - in WhatIf mode, allow processing for simulation
        if ($group.IsDirSynced -eq $true) {
            if (-not $WhatIfMode) {
                Write-Log "  Group '$groupName' is synced from on-premises (IsDirSynced = True). Skipping." -Level Warning
                $stats.GroupsDirSynced++
                continue
            }
            Write-Log "  Group '$groupName' is synced from on-premises (IsDirSynced = True). Processing in WhatIf mode only (no changes will be made)." -Level Warning
        }
        else {
            Write-Log "  Cloud-only group confirmed (IsDirSynced = False)" -Level Debug
        }

        # Get expected members from the Office lookup
        $expectedMembers = @()
        if ($officeLookup.ContainsKey($officeValue)) {
            $expectedMembers = $officeLookup[$officeValue]
        }

        Write-Log "  Expected members based on Office='$officeValue': $($expectedMembers.Count)" -Level Debug

        # Get current group members (UserMailbox only)
        try {
            $currentMembersRaw = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
            $currentMembers = @($currentMembersRaw |
                Where-Object { $_.RecipientType -eq "UserMailbox" } |
                ForEach-Object { $_.PrimarySmtpAddress.ToString().ToLowerInvariant() })

            Write-Log "  Current members (UserMailbox): $($currentMembers.Count)" -Level Debug
        }
        catch {
            Write-Log "  Failed to get members of group '$groupName': $($_.Exception.Message)" -Level Error
            $stats.Errors++
            continue
        }

        # Build hashsets for O(1) comparison
        $expectedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in $expectedMembers) { [void]$expectedSet.Add($m) }

        $currentSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in $currentMembers) { [void]$currentSet.Add($m) }

        # Compute diffs
        $toAdd = [System.Collections.Generic.List[string]]::new()
        foreach ($m in $expectedMembers) {
            if (-not $currentSet.Contains($m)) {
                $toAdd.Add($m)
            }
        }

        $toRemove = [System.Collections.Generic.List[string]]::new()
        foreach ($m in $currentMembers) {
            if (-not $expectedSet.Contains($m)) {
                $toRemove.Add($m)
            }
        }

        if ($toAdd.Count -eq 0 -and $toRemove.Count -eq 0) {
            Write-Log "  Group is already in sync ($($currentMembers.Count) member(s))"
            $stats.MembersAlreadySync++
            continue
        }

        Write-Log "  Changes needed: $($toAdd.Count) to add, $($toRemove.Count) to remove"

        # Add missing members
        foreach ($userSmtp in $toAdd) {
            try {
                if ($WhatIfMode) {
                    Write-Log "  [WHATIF] Would add $userSmtp to $groupName"
                }
                else {
                    Add-DistributionGroupMember -Identity $group.Identity -Member $userSmtp -ErrorAction Stop
                    Write-Log "  [+] Added $userSmtp to $groupName"
                }
                $stats.MembersAdded++
            }
            catch {
                $errMsg = $_.Exception.Message
                if ($errMsg -match "already a member") {
                    Write-Log "  [=] $userSmtp is already a member of $groupName (race condition)" -Level Debug
                    $stats.MembersAlreadySync++
                }
                else {
                    Write-Log "  [!] Failed to add $userSmtp to $groupName - $errMsg" -Level Error
                    $stats.Errors++
                }
            }
        }

        # Remove stale members
        foreach ($userSmtp in $toRemove) {
            try {
                if ($WhatIfMode) {
                    Write-Log "  [WHATIF] Would remove $userSmtp from $groupName"
                }
                else {
                    Remove-DistributionGroupMember -Identity $group.Identity -Member $userSmtp -Confirm:$false -ErrorAction Stop
                    Write-Log "  [-] Removed $userSmtp from $groupName"
                }
                $stats.MembersRemoved++
            }
            catch {
                Write-Log "  [!] Failed to remove $userSmtp from $groupName - $($_.Exception.Message)" -Level Error
                $stats.Errors++
            }
        }
    }

    # -------------------------------------------------------
    # Summary
    # -------------------------------------------------------
    Write-Log ""
    Write-Log "========================================"
    Write-Log "Sync Summary"
    Write-Log "========================================"
    Write-Log "Total mailbox users:           $($stats.TotalUsers)"
    Write-Log "Users skipped (empty Office):  $($stats.UsersSkippedNoOffice)"
    Write-Log "Unique Office values:          $($stats.UniqueOfficeValues)"
    Write-Log "Groups processed:              $($stats.GroupsProcessed)"
    Write-Log "Groups skipped (DirSynced):    $($stats.GroupsDirSynced)"
    Write-Log "Groups skipped (parse error):  $($stats.GroupsSkippedParse)"
    Write-Log "Members added:                 $($stats.MembersAdded)"
    Write-Log "Members removed:               $($stats.MembersRemoved)"
    Write-Log "Groups already in sync:        $($stats.MembersAlreadySync)"
    Write-Log "Errors:                        $($stats.Errors)"
    Write-Log "========================================"
    Write-Log ""
    Write-Log "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Log "========================================"

    # Disconnect
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue

    if ($stats.Errors -gt 0) {
        exit 1
    }
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level Error
    Write-Output "FATAL ERROR: Exception Type: $($_.Exception.GetType().FullName)"
    Write-Output "FATAL ERROR: Stack Trace: $($_.ScriptStackTrace)"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}

#endregion
