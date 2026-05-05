<#
.SYNOPSIS
    Creates dynamic distribution groups in Exchange Online based on a single recipient-filterable attribute.

.DESCRIPTION
    This script creates a set of dynamic distribution groups (DDGs) filtered by a single Exchange
    recipient attribute (e.g., Office, Department, Company, City, Title, CustomAttribute1-15, StateOrProvince).
    
    Each DDG is named using the pattern: <Prefix>_<Value>_<Suffix>@<Domain>
    For example, with -Office D137, the default suffixes produce:
        ACL_D137_BEST@silvan.dk
        ACL_D137_INFO@silvan.dk
        ACL_D137_SALG@silvan.dk

    The DDGs use a RecipientFilter (custom OPATH) scoped to MailboxUsers only.
    After creation, each DDG is verified by previewing its membership using
    Get-Recipient -RecipientPreviewFilter.

    Only ONE filter attribute may be specified per invocation.

    All output is logged to a timestamped log file in the script directory.

.PARAMETER Office
    Filter recipients by the Office attribute.

.PARAMETER Department
    Filter recipients by the Department attribute.

.PARAMETER Company
    Filter recipients by the Company attribute.

.PARAMETER StateOrProvince
    Filter recipients by the StateOrProvince attribute.

.PARAMETER City
    Filter recipients by the City attribute.

.PARAMETER Title
    Filter recipients by the Title attribute.

.PARAMETER CustomAttribute1
    Filter recipients by CustomAttribute1.

.PARAMETER CustomAttribute2
    Filter recipients by CustomAttribute2.

.PARAMETER CustomAttribute3
    Filter recipients by CustomAttribute3.

.PARAMETER CustomAttribute4
    Filter recipients by CustomAttribute4.

.PARAMETER CustomAttribute5
    Filter recipients by CustomAttribute5.

.PARAMETER CustomAttribute6
    Filter recipients by CustomAttribute6.

.PARAMETER CustomAttribute7
    Filter recipients by CustomAttribute7.

.PARAMETER CustomAttribute8
    Filter recipients by CustomAttribute8.

.PARAMETER CustomAttribute9
    Filter recipients by CustomAttribute9.

.PARAMETER CustomAttribute10
    Filter recipients by CustomAttribute10.

.PARAMETER CustomAttribute11
    Filter recipients by CustomAttribute11.

.PARAMETER CustomAttribute12
    Filter recipients by CustomAttribute12.

.PARAMETER CustomAttribute13
    Filter recipients by CustomAttribute13.

.PARAMETER CustomAttribute14
    Filter recipients by CustomAttribute14.

.PARAMETER CustomAttribute15
    Filter recipients by CustomAttribute15.

.PARAMETER ListPrefix
    Prefix used in the DDG name. Default: "ACL"

.PARAMETER Suffixes
    Array of suffixes to create DDGs for. Default: @("BEST","INFO","SALG")

.PARAMETER Domain
    Email domain for the DDG primary SMTP address. Default: "silvan.dk"

.PARAMETER WhatIf
    If set, shows what would be done without making any changes (dry-run mode).

.EXAMPLE
    .\Create-DynamicDistributionLists.ps1 -Office D137
    Creates ACL_D137_BEST@silvan.dk, ACL_D137_INFO@silvan.dk, ACL_D137_SALG@silvan.dk

.EXAMPLE
    .\Create-DynamicDistributionLists.ps1 -Department "Sales" -ListPrefix "DL" -Suffixes @("ALL","MGMT") -Domain "contoso.com"
    Creates DL_Sales_ALL@contoso.com, DL_Sales_MGMT@contoso.com

.EXAMPLE
    .\Create-DynamicDistributionLists.ps1 -Office D137 -WhatIf
    Shows what would be created without making changes.

.NOTES
    Requires: ExchangeOnlineManagement module (automatically installed if missing)
    Authentication: Interactive (browser-based) sign-in to Exchange Online
    
    The script uses the RecipientFilter (custom OPATH) parameter set because many
    filterable properties (e.g., Office, City, Title) are not available as precanned
    Conditional parameters on New-DynamicDistributionGroup.

    Reference: https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-dynamicdistributiongroup
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $false)]
    [string]$Office,

    [Parameter(Mandatory = $false)]
    [string]$Department,

    [Parameter(Mandatory = $false)]
    [string]$Company,

    [Parameter(Mandatory = $false)]
    [string]$StateOrProvince,

    [Parameter(Mandatory = $false)]
    [string]$City,

    [Parameter(Mandatory = $false)]
    [string]$Title,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute1,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute2,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute3,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute4,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute5,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute6,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute7,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute8,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute9,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute10,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute11,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute12,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute13,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute14,

    [Parameter(Mandatory = $false)]
    [string]$CustomAttribute15,

    [Parameter(Mandatory = $false)]
    [string]$ListPrefix = "ACL",

    [Parameter(Mandatory = $false)]
    [string[]]$Suffixes = @("BEST", "INFO", "SALG"),

    [Parameter(Mandatory = $false)]
    [string]$Domain = "silvan.dk"
)

$ErrorActionPreference = "Stop"

#region Log File Initialization

$scriptPath = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptPath)) {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path $scriptPath "Create-DynamicDistributionLists_$timestamp.log"

#endregion

#region Functions

function Write-Log {
    <#
    .SYNOPSIS
        Writes a message to the console and log file with level-based coloring.
    #>
    param (
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info'    { 'White' }
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }

    # Write to console
    Write-Host "[$ts] [$Level] $Message" -ForegroundColor $color

    # Write to log file with retry logic
    $logEntry = "[$ts] [$Level] $Message"
    $maxRetries = 3
    $retryCount = 0
    $success = $false

    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction Stop
            $success = $true
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Start-Sleep -Seconds ([Math]::Pow(2, $retryCount - 1))
            }
            else {
                Write-Host "[WARNING] Failed to write to log file after $maxRetries attempts: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}

function Test-Administrator {
    <#
    .SYNOPSIS
        Checks if the current session is running as administrator.
    #>
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Install-ExchangeOnlineModule {
    <#
    .SYNOPSIS
        Checks for ExchangeOnlineManagement module and installs it if missing.
    #>
    Write-Log "Checking for ExchangeOnlineManagement module..." -Level Info

    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

    if ($module) {
        Write-Log "Found ExchangeOnlineManagement module version $($module.Version)" -Level Success
        try {
            Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            Write-Log "Module imported successfully" -Level Success
            return $true
        }
        catch {
            Write-Log "Failed to import ExchangeOnlineManagement module: $($_.Exception.Message)" -Level Error
            return $false
        }
    }

    Write-Log "ExchangeOnlineManagement module not found. Installing..." -Level Warning

    $isAdmin = Test-Administrator
    $scope = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }
    Write-Log "Installing for scope: $scope" -Level Info

    try {
        Install-Module -Name ExchangeOnlineManagement -Scope $scope -Force -AllowClobber -ErrorAction Stop
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        Write-Log "ExchangeOnlineManagement module installed and imported successfully" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Connect-ToExchangeOnline {
    <#
    .SYNOPSIS
        Connects to Exchange Online, checking for an existing connection first.
    #>
    Write-Log "Checking Exchange Online connection..." -Level Info

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Already connected to Exchange Online" -Level Success
        return $true
    }
    catch {
        Write-Log "Not connected to Exchange Online. Connecting..." -Level Info
    }

    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Test-DynamicDistributionGroup {
    <#
    .SYNOPSIS
        Verifies a dynamic distribution group by previewing its membership.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    Write-Log "Verifying DDG '$Identity'..." -Level Info

    try {
        $ddg = Get-DynamicDistributionGroup -Identity $Identity -ErrorAction Stop
        $filter = $ddg.RecipientFilter
        Write-Log "  RecipientFilter: $filter" -Level Info

        $recipients = @(Get-Recipient -RecipientPreviewFilter $filter -ErrorAction Stop)
        Write-Log "  Matched recipients: $($recipients.Count)" -Level Success

        if ($recipients.Count -gt 0) {
            foreach ($recipient in $recipients) {
                Write-Log "    - $($recipient.DisplayName) ($($recipient.PrimarySmtpAddress))" -Level Info
            }
        }
        else {
            Write-Log "  No recipients matched the filter. This may be expected if no mailbox users match the criteria yet." -Level Warning
        }

        return $recipients.Count
    }
    catch {
        Write-Log "  Verification failed for '$Identity': $($_.Exception.Message)" -Level Error
        return -1
    }
}

#endregion

#region Parameter Validation

# Map of filter parameter names to their OPATH RecipientFilter property names
$filterParameters = @{
    'Office'           = 'Office'
    'Department'       = 'Department'
    'Company'          = 'Company'
    'StateOrProvince'  = 'StateOrProvince'
    'City'             = 'City'
    'Title'            = 'Title'
    'CustomAttribute1' = 'CustomAttribute1'
    'CustomAttribute2' = 'CustomAttribute2'
    'CustomAttribute3' = 'CustomAttribute3'
    'CustomAttribute4' = 'CustomAttribute4'
    'CustomAttribute5' = 'CustomAttribute5'
    'CustomAttribute6' = 'CustomAttribute6'
    'CustomAttribute7' = 'CustomAttribute7'
    'CustomAttribute8' = 'CustomAttribute8'
    'CustomAttribute9' = 'CustomAttribute9'
    'CustomAttribute10' = 'CustomAttribute10'
    'CustomAttribute11' = 'CustomAttribute11'
    'CustomAttribute12' = 'CustomAttribute12'
    'CustomAttribute13' = 'CustomAttribute13'
    'CustomAttribute14' = 'CustomAttribute14'
    'CustomAttribute15' = 'CustomAttribute15'
}

# Determine which filter parameters were specified
$specifiedFilters = @()
foreach ($paramName in $filterParameters.Keys) {
    $value = (Get-Variable -Name $paramName -ErrorAction SilentlyContinue).Value
    if (-not [string]::IsNullOrWhiteSpace($value)) {
        $specifiedFilters += @{ Name = $paramName; Value = $value; OPathProperty = $filterParameters[$paramName] }
    }
}

if ($specifiedFilters.Count -eq 0) {
    Write-Log "ERROR: You must specify exactly one filter parameter (e.g., -Office, -Department, -Company, -City, -Title, -StateOrProvince, -CustomAttribute1 through -CustomAttribute15)." -Level Error
    exit 1
}

if ($specifiedFilters.Count -gt 1) {
    $names = ($specifiedFilters | ForEach-Object { "-$($_.Name)" }) -join ", "
    Write-Log "ERROR: Only one filter parameter may be specified per invocation. You specified: $names" -Level Error
    exit 1
}

$chosenFilter = $specifiedFilters[0]
$filterPropertyName = $chosenFilter.OPathProperty
$filterValue = $chosenFilter.Value

#endregion

#region Main Script

Write-Log "==========================================" -Level Info
Write-Log "Create Dynamic Distribution Lists" -Level Info
Write-Log "==========================================" -Level Info
Write-Log "Filter: $filterPropertyName -eq '$filterValue'" -Level Info
Write-Log "Prefix: $ListPrefix" -Level Info
Write-Log "Suffixes: $($Suffixes -join ', ')" -Level Info
Write-Log "Domain: $Domain" -Level Info
Write-Log "WhatIf: $($WhatIfPreference)" -Level Info
Write-Log "Log File: $LogFile" -Level Info
Write-Log "==========================================" -Level Info

# Build the OPATH RecipientFilter (scoped to MailboxUsers)
$recipientFilter = "((RecipientType -eq 'UserMailbox') -and ($filterPropertyName -eq '$filterValue'))"
Write-Log "RecipientFilter: $recipientFilter" -Level Info

# Install module and connect
if (-not (Install-ExchangeOnlineModule)) {
    Write-Log "Cannot proceed without ExchangeOnlineManagement module." -Level Error
    exit 1
}

if (-not (Connect-ToExchangeOnline)) {
    Write-Log "Cannot proceed without Exchange Online connection." -Level Error
    exit 1
}

# Statistics
$stats = @{
    Total     = $Suffixes.Count
    Created   = 0
    Skipped   = 0
    Failed    = 0
    Verified  = 0
}

# Create DDGs
foreach ($suffix in $Suffixes) {
    $ddgName = "${ListPrefix}_${filterValue}_${suffix}"
    $ddgEmail = "${ddgName}@${Domain}"

    Write-Log "" -Level Info
    Write-Log "Processing: $ddgName ($ddgEmail)" -Level Info

    # Check if DDG already exists
    try {
        $existing = Get-DynamicDistributionGroup -Identity $ddgName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "DDG '$ddgName' already exists. Skipping creation." -Level Warning
            $stats.Skipped++

            # Still verify the existing DDG
            Write-Log "Waiting 10 seconds before verification..." -Level Info
            Start-Sleep -Seconds 10
            $recipientCount = Test-DynamicDistributionGroup -Identity $ddgName
            if ($recipientCount -ge 0) { $stats.Verified++ }
            continue
        }
    }
    catch {
        # Get-DynamicDistributionGroup throws if not found with -ErrorAction SilentlyContinue in some cases
        # This is expected, proceed with creation
    }

    # Create DDG
    if ($PSCmdlet.ShouldProcess($ddgEmail, "Create Dynamic Distribution Group")) {
        try {
            Write-Log "Creating DDG '$ddgName' with PrimarySmtpAddress '$ddgEmail'..." -Level Info
            New-DynamicDistributionGroup `
                -Name $ddgName `
                -DisplayName $ddgName `
                -RecipientFilter $recipientFilter `
                -PrimarySmtpAddress $ddgEmail `
                -ErrorAction Stop

            Write-Log "Successfully created DDG '$ddgName'" -Level Success
            $stats.Created++

            # Wait before verification
            Write-Log "Waiting 10 seconds before verification..." -Level Info
            Start-Sleep -Seconds 10

            # Verify
            $recipientCount = Test-DynamicDistributionGroup -Identity $ddgName
            if ($recipientCount -ge 0) { $stats.Verified++ }
        }
        catch {
            Write-Log "Failed to create DDG '$ddgName': $($_.Exception.Message)" -Level Error
            $stats.Failed++
        }
    }
    else {
        Write-Log "WHATIF: Would create DDG '$ddgName' with:" -Level Info
        Write-Log "  PrimarySmtpAddress: $ddgEmail" -Level Info
        Write-Log "  RecipientFilter: $recipientFilter" -Level Info
        $stats.Created++
    }
}

# Summary
Write-Log "" -Level Info
Write-Log "==========================================" -Level Info
Write-Log "Summary" -Level Info
Write-Log "==========================================" -Level Info
Write-Log "Total suffixes: $($stats.Total)" -Level Info
Write-Log "DDGs created: $($stats.Created)" -Level Success
Write-Log "DDGs skipped (already existed): $($stats.Skipped)" -Level Warning
Write-Log "DDGs failed: $($stats.Failed)" -Level $(if ($stats.Failed -gt 0) { 'Error' } else { 'Info' })
Write-Log "DDGs verified: $($stats.Verified)" -Level Info
Write-Log "==========================================" -Level Info

if ($stats.Failed -gt 0) {
    Write-Log "WARNING: Some DDGs failed to create. Review the log above for details." -Level Warning
}

Write-Log "Log file: $LogFile" -Level Info
Write-Log "Script completed." -Level Success

# Disconnect
try {
    Write-Log "Disconnecting from Exchange Online..." -Level Info
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Disconnected successfully" -Level Success
}
catch {
    Write-Log "WARNING: Failed to disconnect cleanly: $($_.Exception.Message)" -Level Warning
}

#endregion
