<#
.SYNOPSIS
    Enables archive mailboxes for Exchange Online users from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing user principal names (UPN) and enables
    the archive mailbox feature for each user in Exchange Online.

.PARAMETER CsvFile
    Path to the CSV file containing mailbox identities. The CSV must have a column header "UPN".

.PARAMETER LogFile
    Path to log file. If not specified, creates a timestamped log in the script directory.

.EXAMPLE
    .\Set-MailboxArchive.ps1 -CsvFile "C:\Users\mailboxes.csv"

.EXAMPLE
    .\Set-MailboxArchive.ps1 -CsvFile ".\mailboxes.csv" -LogFile "C:\Logs\ArchiveActivation.log"

.NOTES
    Prerequisites:
    - Exchange Online Management module (automatically installed if missing)
    - Must be connected to Exchange Online (automatically connects)
    - Requires appropriate Exchange Online administrator permissions
    - Users must have a license that supports archive mailboxes
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file containing UPN column")]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { throw "CSV file not found: $_" }
    })]
    [string]$CsvFile,

    [Parameter(Mandatory = $false, HelpMessage = "Path to log file. If not specified, creates log in script directory")]
    [string]$LogFile
)

# Initialize log file path
if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile = Join-Path $scriptPath "MailboxArchive_$timestamp.log"
}

# Ensure log directory exists
$logDirectory = Split-Path -Parent $LogFile
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
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
Write-Log "=== Starting Mailbox Archive Activation ===" -Level Info
Write-Log "CSV File: $CsvFile" -Level Info
Write-Log "Log File: $LogFile" -Level Info
Write-Log "" -Level Info

# Ensure Exchange Online Management module is installed
if (-not (Install-ExchangeOnlineModule)) {
    Write-Log "Cannot proceed without Exchange Online Management module" -Level Error
    exit 1
}

# Connect to Exchange Online
if (-not (Connect-ToExchangeOnline)) {
    Write-Log "Cannot proceed without Exchange Online connection" -Level Error
    exit 1
}

# Import CSV file
try {
    $mailboxes = Import-Csv -Path $CsvFile -ErrorAction Stop -Delimiter ","
    Write-Log "Successfully imported $($mailboxes.Count) mailbox(es) from CSV" -Level Success
}
catch {
    Write-Log "Failed to import CSV file: $($_.Exception.Message)" -Level Error
    exit 1
}

# Initialize counters
$successCount = 0
$failureCount = 0
$skippedCount = 0
$totalCount = $mailboxes.Count

# Process each mailbox
Write-Log "Processing mailboxes..." -Level Info
foreach ($mailbox in $mailboxes) {
    
    if ([string]::IsNullOrWhiteSpace($mailbox.UPN)) {
        Write-Log "Skipping empty UPN entry" -Level Warning
        $skippedCount++
        continue
    }
    
    try {
        # Check if archive is already enabled
        $mbx = Get-Mailbox -Identity $mailbox.UPN -ErrorAction Stop
        
        if ($mbx.ArchiveStatus -eq 'Active') {
            Write-Log "SKIPPED: $($mailbox.UPN) - Archive already enabled" -Level Warning
            $skippedCount++
            continue
        }
        
        # Enable archive
        Enable-Mailbox -Identity $mailbox.UPN -Archive -ErrorAction Stop
        Write-Log "SUCCESS: $($mailbox.UPN) - Archive enabled" -Level Success
        $successCount++
    }
    catch {
        Write-Log "FAILED: $($mailbox.UPN) - $($_.Exception.Message)" -Level Error
        $failureCount++
    }
}

# Summary
Write-Log "" -Level Info
Write-Log "=== Processing Complete ===" -Level Info
Write-Log "Total mailboxes: $totalCount" -Level Info
Write-Log "Successful: $successCount" -Level Success
Write-Log "Skipped (already enabled): $skippedCount" -Level Warning
Write-Log "Failed: $failureCount" -Level $(if ($failureCount -gt 0) { 'Error' } else { 'Info' })
