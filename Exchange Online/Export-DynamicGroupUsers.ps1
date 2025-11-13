<#
.SYNOPSIS
    Exports users from a Microsoft 365 dynamic group who have Exchange Online licenses.

.DESCRIPTION
    This script retrieves members of a specified dynamic group and filters for users with
    Exchange Online Plan 1, Exchange Online Plan 2, or Microsoft 365 E5 licenses.
    Exports the User Principal Names to a CSV file.

.PARAMETER DynamicGroupName
    The display name or object ID of the dynamic group to export users from.

.PARAMETER OutputCsvFile
    Path to the output CSV file. If not specified, creates a timestamped CSV in the script directory.

.PARAMETER LogFile
    Path to log file. If not specified, creates a timestamped log in the script directory.

.EXAMPLE
    .\Export-DynamicGroupUsers.ps1 -DynamicGroupName "All Licensed Users"

.EXAMPLE
    .\Export-DynamicGroupUsers.ps1 -DynamicGroupName "Sales Team" -OutputCsvFile "C:\Exports\SalesUsers.csv"

.NOTES
    Prerequisites:
    - Microsoft.Graph PowerShell module (automatically installed if missing)
    - Exchange Online Management module (automatically installed if missing)
    - Requires appropriate Microsoft 365 administrator permissions (User.Read.All, Group.Read.All)
    
    Supported License SKUs:
    - Exchange Online (Plan 1)
    - Exchange Online (Plan 2)
    - Microsoft 365 E5
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Display name or Object ID of the dynamic group")]
    [ValidateNotNullOrEmpty()]
    [string]$DynamicGroupName,

    [Parameter(Mandatory = $false, HelpMessage = "Path to output CSV file")]
    [string]$OutputCsvFile,

    [Parameter(Mandatory = $false, HelpMessage = "Path to log file")]
    [string]$LogFile
)

# Sanitize group name for filename (remove invalid characters)
$sanitizedGroupName = $DynamicGroupName -replace '[\\/:*?"<>|]', '_'

# Initialize log file path
if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile = Join-Path $scriptPath "$sanitizedGroupName`_$timestamp.log"
}

# Initialize output CSV file path
if ([string]::IsNullOrWhiteSpace($OutputCsvFile)) {
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputCsvFile = Join-Path $scriptPath "$sanitizedGroupName`_$timestamp.csv"
}

# Ensure log directory exists
$logDirectory = Split-Path -Parent $LogFile
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# Ensure output directory exists
$outputDirectory = Split-Path -Parent $OutputCsvFile
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

# Function to write log messages
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        'Info'    { 'White' }
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
    }
    
    # Write to console
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
    
    # Write to log file
    $logEntry = "[$timestamp] [$Level] $Message"
    try {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Host "[WARNING] Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to ensure Microsoft Graph module is installed
function Install-GraphModule {
    Write-Log "Checking Microsoft Graph PowerShell module..." -Level Info
    
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) {
        Write-Log "Microsoft Graph module is already installed" -Level Success
        return $true
    }
    
    Write-Log "Microsoft Graph module not found. Installing..." -Level Warning
    
    # Determine installation scope based on admin context
    $isAdmin = Test-Administrator
    $scope = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }
    
    Write-Log "Installing for scope: $scope" -Level Info
    
    try {
        Install-Module -Name Microsoft.Graph.Users -Scope $scope -Force -AllowClobber -ErrorAction Stop
        Install-Module -Name Microsoft.Graph.Groups -Scope $scope -Force -AllowClobber -ErrorAction Stop
        Write-Log "Microsoft Graph module installed successfully" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to install Microsoft Graph module: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to ensure Exchange Online Management module is installed
function Install-ExchangeOnlineModule {
    Write-Log "Checking Exchange Online Management module..." -Level Info
    
    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Log "Exchange Online Management module is already installed" -Level Success
        return $true
    }
    
    Write-Log "Exchange Online Management module not found. Installing..." -Level Warning
    
    # Determine installation scope based on admin context
    $isAdmin = Test-Administrator
    $scope = if ($isAdmin) { "AllUsers" } else { "CurrentUser" }
    
    Write-Log "Installing for scope: $scope" -Level Info
    
    try {
        Install-Module -Name ExchangeOnlineManagement -Scope $scope -Force -AllowClobber -ErrorAction Stop
        Write-Log "Exchange Online Management module installed successfully" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to install Exchange Online Management module: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to connect to Microsoft Graph
function Connect-ToMicrosoftGraph {
    Write-Log "Checking Microsoft Graph connection..." -Level Info
    
    # Check if already connected
    try {
        $context = Get-MgContext -ErrorAction Stop
        if ($null -ne $context) {
            Write-Log "Already connected to Microsoft Graph" -Level Success
            return $true
        }
    }
    catch {
        Write-Log "Not connected to Microsoft Graph. Connecting..." -Level Warning
    }
    
    # Attempt to connect
    try {
        Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All" -NoWelcome -ErrorAction Stop
        Write-Log "Successfully connected to Microsoft Graph" -Level Success
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Function to connect to Exchange Online
function Connect-ToExchangeOnline {
    Write-Log "Checking Exchange Online connection..." -Level Info
    
    # Check if already connected
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Already connected to Exchange Online" -Level Success
        return $true
    }
    catch {
        Write-Log "Not connected to Exchange Online. Connecting..." -Level Warning
    }
    
    # Attempt to connect
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

# Start script
Write-Log "=== Starting Dynamic Group User Export ===" -Level Info
Write-Log "Dynamic Group: $DynamicGroupName" -Level Info
Write-Log "Output CSV: $OutputCsvFile" -Level Info
Write-Log "Log File: $LogFile" -Level Info
Write-Log "" -Level Info

# Ensure modules are installed
if (-not (Install-GraphModule)) {
    Write-Log "Cannot proceed without Microsoft Graph module" -Level Error
    exit 1
}

if (-not (Install-ExchangeOnlineModule)) {
    Write-Log "Cannot proceed without Exchange Online Management module" -Level Error
    exit 1
}

# Connect to services
if (-not (Connect-ToMicrosoftGraph)) {
    Write-Log "Cannot proceed without Microsoft Graph connection" -Level Error
    exit 1
}

if (-not (Connect-ToExchangeOnline)) {
    Write-Log "Cannot proceed without Exchange Online connection" -Level Error
    exit 1
}

# Find the dynamic group
Write-Log "Searching for dynamic group..." -Level Info
try {
    $group = Get-MgGroup -Filter "displayName eq '$DynamicGroupName'" -ErrorAction Stop | Select-Object -First 1
    
    if (-not $group) {
        # Try by Object ID
        $group = Get-MgGroup -GroupId $DynamicGroupName -ErrorAction Stop
    }
    
    if (-not $group) {
        Write-Log "Dynamic group '$DynamicGroupName' not found" -Level Error
        exit 1
    }
    
    Write-Log "Found group: $($group.DisplayName) (ID: $($group.Id))" -Level Success
    
    # Verify it's a dynamic group
    if ($group.GroupTypes -notcontains "DynamicMembership") {
        Write-Log "Warning: '$($group.DisplayName)' is not a dynamic group" -Level Warning
    }
}
catch {
    Write-Log "Failed to find group: $($_.Exception.Message)" -Level Error
    exit 1
}

# Define Exchange Online license SKU part numbers
$exchangeLicenseSKUs = @(
    "EXCHANGESTANDARD",      # Exchange Online (Plan 1)
    "EXCHANGEENTERPRISE",    # Exchange Online (Plan 2)
    "SPE_E5",                # Microsoft 365 E5
    "ENTERPRISEPREMIUM",     # Microsoft 365 E5 (alternative)
    "SPE_E3",                # Microsoft 365 E3 (includes Exchange)
    "ENTERPRISEPACK",        # Office 365 E3
    "STANDARDPACK",          # Office 365 E1
    "EXCHANGEDESKLESS"       # Exchange Online Kiosk
)

Write-Log "Retrieving group members..." -Level Info

# Get all members of the group
try {
    $groupMembers = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
    Write-Log "Retrieved $($groupMembers.Count) members from group" -Level Success
}
catch {
    Write-Log "Failed to retrieve group members: $($_.Exception.Message)" -Level Error
    exit 1
}

# Initialize results array
$qualifiedUsers = @()
$processedCount = 0
$skippedCount = 0

Write-Log "Processing group members..." -Level Info

foreach ($member in $groupMembers) {
    $processedCount++
    
    # Only process user objects
    if ($member.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.user') {
        Write-Log "Skipping non-user member: $($member.Id)" -Level Warning
        $skippedCount++
        continue
    }
    
    try {
        # Get full user details
        $user = Get-MgUser -UserId $member.Id -Property "UserPrincipalName,DisplayName" -ErrorAction Stop
        
        # Check if user has any of the Exchange Online licenses
        $hasExchangeLicense = $false
        $userLicenses = @()
        
        # Get license details using dedicated cmdlet
        try {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $member.Id -ErrorAction Stop
            
            if ($licenseDetails -and $licenseDetails.Count -gt 0) {
                foreach ($licenseDetail in $licenseDetails) {
                    $skuPartNumber = $licenseDetail.SkuPartNumber
                    $userLicenses += $skuPartNumber
                    
                    # Check if SKU part number matches our Exchange licenses
                    if ($exchangeLicenseSKUs -contains $skuPartNumber) {
                        $hasExchangeLicense = $true
                        Write-Log "  Found qualifying license for $($user.UserPrincipalName): $skuPartNumber" -Level Info
                        break
                    }
                }
            }
            else {
                Write-Log "  User $($user.UserPrincipalName) has no licenses" -Level Warning
            }
            
            # Log all licenses if no match found (for debugging)
            if (-not $hasExchangeLicense -and $userLicenses.Count -gt 0) {
                Write-Log "  User $($user.UserPrincipalName) has licenses: $($userLicenses -join ', ') - None match Exchange criteria" -Level Warning
            }
        }
        catch {
            Write-Log "  Could not retrieve licenses for $($user.UserPrincipalName): $($_.Exception.Message)" -Level Warning
        }
        
        if ($hasExchangeLicense) {
            # Verify user has a mailbox using Get-EXOMailbox (fast with Minimum property set)
            try {
                $mailbox = Get-EXOMailbox -Identity $user.UserPrincipalName -PropertySets Minimum -ErrorAction Stop
                
                if ($mailbox) {
                    $qualifiedUsers += [PSCustomObject]@{
                        UPN = $user.UserPrincipalName
                    }
                    Write-Log "Added: $($user.UserPrincipalName) [Type: $($mailbox.RecipientTypeDetails)]" -Level Success
                }
            }
            catch {
                Write-Log "Skipped: $($user.UserPrincipalName) - No mailbox found: $($_.Exception.Message)" -Level Warning
                $skippedCount++
            }
        }
        else {
            $skippedCount++
        }
    }
    catch {
        Write-Log "Error processing member $($member.Id): $($_.Exception.Message)" -Level Error
        $skippedCount++
    }
    
    # Progress indicator every 10 users
    if ($processedCount % 10 -eq 0) {
        Write-Log "Progress: Processed $processedCount of $($groupMembers.Count) members..." -Level Info
    }
}

# Export to CSV
Write-Log "" -Level Info
if ($qualifiedUsers.Count -gt 0) {
    try {
        $qualifiedUsers | Export-Csv -Path $OutputCsvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "Successfully exported $($qualifiedUsers.Count) users to: $OutputCsvFile" -Level Success
    }
    catch {
        Write-Log "Failed to export CSV: $($_.Exception.Message)" -Level Error
        exit 1
    }
}
else {
    Write-Log "No qualified users found to export" -Level Warning
}

# Summary
Write-Log "" -Level Info
Write-Log "=== Processing Complete ===" -Level Info
Write-Log "Total members processed: $processedCount" -Level Info
Write-Log "Qualified users exported: $($qualifiedUsers.Count)" -Level Success
Write-Log "Skipped: $skippedCount" -Level Warning
Write-Log "Output file: $OutputCsvFile" -Level Info
