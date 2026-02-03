#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Verifies archive and retention policy settings for mailboxes from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing email addresses and retrieves
    mailbox information including display name, primary email, retention policy,
    and archive status. Results are displayed in Out-GridView.

.PARAMETER CsvFile
    Path to the CSV file containing email addresses. The CSV must have a column with email addresses.

.PARAMETER EmailColumn
    Name of the column containing email addresses. Default is "Email".

.PARAMETER OutputFile
    Path to the output CSV file. If not specified, creates a timestamped file in the script directory.

.EXAMPLE
    .\verify-archive-retention.ps1 -CsvFile "C:\mailboxes.csv"

.EXAMPLE
    .\verify-archive-retention.ps1 -CsvFile "C:\mailboxes.csv" -EmailColumn "PrimarySmtpAddress"

.EXAMPLE
    .\verify-archive-retention.ps1 -CsvFile "C:\mailboxes.csv" -OutputFile "C:\Results\verification.csv"

.NOTES
    Prerequisites:
    - Exchange Online Management module
    - Must be connected to Exchange Online or will prompt for connection
    - Requires appropriate Exchange Online administrator permissions
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file containing email addresses")]
    [ValidateScript({
        if (Test-Path $_) { $true }
        else { throw "CSV file not found: $_" }
    })]
    [string]$CsvFile,

    [Parameter(Mandatory = $false, HelpMessage = "Name of the column containing email addresses")]
    [string]$EmailColumn = "Email",

    [Parameter(Mandatory = $false, HelpMessage = "Path to output CSV file. If not specified, creates timestamped file in script directory")]
    [string]$OutputFile
)

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
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Connect-ToExchangeOnline {
    Write-Log "Checking Exchange Online connection..." -Level Info
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Already connected to Exchange Online" -Level Success
        return $true
    }
    catch {
        Write-Log "Not connected to Exchange Online. Connecting..." -Level Warning
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

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFile = Join-Path $scriptPath "ArchiveRetentionVerification_$timestamp.csv"
}

$outputDirectory = Split-Path -Parent $OutputFile
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}

Write-Log "=== Mailbox Archive and Retention Verification ===" -Level Info
Write-Log "CSV File: $CsvFile" -Level Info
Write-Log "Email Column: $EmailColumn" -Level Info
Write-Log "Output File: $OutputFile" -Level Info
Write-Log "" -Level Info

if (-not (Connect-ToExchangeOnline)) {
    Write-Log "Cannot proceed without Exchange Online connection" -Level Error
    exit 1
}

try {
    $csvData = Import-Csv -Path $CsvFile -ErrorAction Stop
    Write-Log "Successfully imported $($csvData.Count) row(s) from CSV" -Level Success
}
catch {
    Write-Log "Failed to import CSV file: $($_.Exception.Message)" -Level Error
    exit 1
}

if ($EmailColumn -notin $csvData[0].PSObject.Properties.Name) {
    Write-Log "Column '$EmailColumn' not found in CSV. Available columns: $($csvData[0].PSObject.Properties.Name -join ', ')" -Level Error
    exit 1
}

$results = @()
$processedCount = 0
$totalCount = $csvData.Count

Write-Log "Processing mailboxes..." -Level Info

foreach ($row in $csvData) {
    $processedCount++
    $emailAddress = $row.$EmailColumn
    
    if ([string]::IsNullOrWhiteSpace($emailAddress)) {
        Write-Log "Skipping empty email address at row $processedCount" -Level Warning
        continue
    }
    
    Write-Progress -Activity "Verifying Mailboxes" -Status "Processing $emailAddress ($processedCount of $totalCount)" -PercentComplete (($processedCount / $totalCount) * 100)
    
    try {
        $mailbox = Get-Mailbox -Identity $emailAddress -ErrorAction Stop
        
        $result = [PSCustomObject]@{
            DisplayName       = $mailbox.DisplayName
            PrimaryEmail      = $mailbox.PrimarySmtpAddress
            RetentionPolicy   = if ($mailbox.RetentionPolicy) { $mailbox.RetentionPolicy } else { "None" }
            ArchiveEnabled    = $mailbox.ArchiveStatus -eq 'Active'
            ArchiveStatus     = $mailbox.ArchiveStatus
            ArchiveDatabase   = if ($mailbox.ArchiveDatabase) { $mailbox.ArchiveDatabase } else { "N/A" }
            MailboxType       = $mailbox.RecipientTypeDetails
        }
        
        $results += $result
        Write-Log "SUCCESS: $emailAddress - Archive: $($result.ArchiveEnabled), Policy: $($result.RetentionPolicy)" -Level Success
    }
    catch {
        Write-Log "FAILED: $emailAddress - $($_.Exception.Message)" -Level Error
        
        $result = [PSCustomObject]@{
            DisplayName       = "ERROR"
            PrimaryEmail      = $emailAddress
            RetentionPolicy   = "Error retrieving"
            ArchiveEnabled    = $false
            ArchiveStatus     = "Error"
            ArchiveDatabase   = "N/A"
            MailboxType       = "Unknown"
        }
        $results += $result
    }
}

Write-Progress -Activity "Verifying Mailboxes" -Completed

Write-Log "" -Level Info
Write-Log "=== Processing Complete ===" -Level Info
Write-Log "Total processed: $processedCount" -Level Info
Write-Log "Successful: $($results | Where-Object { $_.DisplayName -ne 'ERROR' } | Measure-Object | Select-Object -ExpandProperty Count)" -Level Success
Write-Log "Failed: $($results | Where-Object { $_.DisplayName -eq 'ERROR' } | Measure-Object | Select-Object -ExpandProperty Count)" -Level Error
Write-Log "" -Level Info

if ($results.Count -gt 0) {
    try {
        Write-Log "Exporting results to CSV file..." -Level Info
        $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "Results exported successfully to: $OutputFile" -Level Success
        Write-Log "Total records exported: $($results.Count)" -Level Info
    }
    catch {
        Write-Log "Failed to export results: $($_.Exception.Message)" -Level Error
        exit 1
    }
}
else {
    Write-Log "No results to export" -Level Warning
}

Write-Log "Script completed successfully" -Level Success
