param(
    [Parameter(Mandatory = $true)]
    [string]
    $CsvFile,

    [Parameter(Mandatory = $false)]
    [string]
    $LogFile
)

if ([string]::IsNullOrWhiteSpace($LogFile)) {
    $scriptPath = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile = Join-Path $scriptPath ("ManagedFolderAssistant_{0}.log" -f $timestamp)
}

$logDirectory = Split-Path -Parent $LogFile
if (-not (Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

$script:LogFile = $LogFile

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

    $logEntry = "[$timestamp] [$Level] $Message"
    try {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction Stop
    }
    catch {
        Write-Host "[WARNING] Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Log "=== Starting Managed Folder Assistant run ===" -Level Info
Write-Log "Input CSV: $CsvFile" -Level Info
Write-Log "Log File: $LogFile" -Level Info

try {
    Write-Log "Checking Exchange Online Management module..." -Level Info
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Log "ExchangeOnlineManagement module is not installed. Install-Module ExchangeOnlineManagement -Scope CurrentUser" -Level Error
        throw "ExchangeOnlineManagement module is not installed. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Write-Log "Checking Exchange Online connection..." -Level Info
    $exoSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $exoSession) {
        Write-Log "Not connected to Exchange Online. Connecting..." -Level Warning
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Log "Successfully connected to Exchange Online" -Level Success
    }
    else {
        Write-Log "Already connected to Exchange Online" -Level Success
    }
}
catch {
    Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level Error
    exit 1
}

if (-not (Test-Path -LiteralPath $CsvFile)) {
    Write-Log "Input CSV path not found: $CsvFile" -Level Error
    exit 1
}

try {
    $inputData = Import-Csv -LiteralPath $CsvFile
    Write-Log "Successfully imported $($inputData.Count) row(s) from CSV" -Level Success
}
catch {
    Write-Log "Failed to import CSV file: $($_.Exception.Message)" -Level Error
    exit 1
}

if (-not $inputData) {
    Write-Log "Input CSV is empty or could not be parsed: $CsvFile" -Level Error
    exit 1
}

$results = @()
$successCount = 0
$failureCount = 0
$totalCount = $inputData.Count

Write-Log "Processing mailboxes..." -Level Info

foreach ($row in $inputData) {
    $upn = $null
    if ($row.PSObject.Properties.Name -contains 'UPN') {
        $upn = $row.UPN
    } elseif ($row.PSObject.Properties.Name -contains 'UserPrincipalName') {
        $upn = $row.UserPrincipalName
    }

    if (-not $upn) {
        Write-Log "Skipping row with missing UPN/UserPrincipalName" -Level Warning
        $results += [pscustomobject]@{
            UPN        = $null
            Status     = 'Failed'
            Error      = 'Missing UPN/UserPrincipalName column value in row'
            TimeStamp  = [DateTime]::UtcNow
        }
        $failureCount++
        continue
    }

    try {
        Start-ManagedFolderAssistant -Identity $upn -ErrorAction Stop
        Write-Log "SUCCESS: $upn - Managed Folder Assistant started" -Level Success

        $results += [pscustomobject]@{
            UPN        = $upn
            Status     = 'Success'
            Error      = $null
            TimeStamp  = [DateTime]::UtcNow
        }
        $successCount++
    }
    catch {
        Write-Log "FAILED: $upn - $($_.Exception.Message)" -Level Error
        $results += [pscustomobject]@{
            UPN        = $upn
            Status     = 'Failed'
            Error      = $_.Exception.Message
            TimeStamp  = [DateTime]::UtcNow
        }
        $failureCount++
    }

    # Wait 5 seconds before processing the next mailbox
    Start-Sleep -Seconds 5
}
Write-Log "=== Processing Complete ===" -Level Info
Write-Log "Total mailboxes: $totalCount" -Level Info
Write-Log "Successful: $successCount" -Level Success
Write-Log "Failed: $failureCount" -Level $(if ($failureCount -gt 0) { 'Warning' } else { 'Info' })
