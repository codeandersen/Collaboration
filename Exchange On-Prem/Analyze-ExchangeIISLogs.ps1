#Requires -Version 5.1
<#
.SYNOPSIS
    Exchange IIS Log Analyzer for troubleshooting connection slowness.

.DESCRIPTION
    Parses IIS logs from Exchange 2016 servers and analyzes performance metrics for all major endpoints:
    - /mapi/ (MAPI over HTTP)
    - /Microsoft-Server-ActiveSync (ActiveSync)
    - /EWS (Exchange Web Services)
    - /owa/ (Outlook Web App)
    - /Autodiscover/
    - /PowerShell/
    - /OAB/ (Offline Address Book)
    
    Generates console output, HTML reports, and CSV exports with response time statistics,
    error rates, and request volume analysis.

.PARAMETER LogPath
    Path to IIS log directory. Default: C:\inetpub\logs\LogFiles

.PARAMETER OutputPath
    Directory for HTML and CSV output files. Default: Script directory.

.PARAMETER StartDate
    Start date for log analysis (inclusive). Format: yyyy-MM-dd

.PARAMETER EndDate
    End date for log analysis (inclusive). Format: yyyy-MM-dd

.PARAMETER TopN
    Number of top slow requests to display. Default: 20

.PARAMETER IncludeAllRequests
    Include all requests in analysis, not just Exchange endpoints.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -LogPath "C:\inetpub\logs\LogFiles\W3SVC1"
    Analyzes all logs in the specified directory.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -LogPath "C:\inetpub\logs\LogFiles\W3SVC1" -StartDate "2026-05-01" -EndDate "2026-05-05"
    Analyzes logs within the specified date range.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -LogPath "C:\inetpub\logs\LogFiles\W3SVC1" -OutputPath "C:\Reports" -TopN 50
    Outputs reports to C:\Reports and shows top 50 slow requests.

.NOTES
    Author: Exchange IIS Log Analyzer
    Version: 1.0
    Requires: Read access to IIS log files
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\inetpub\logs\LogFiles",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = $PSScriptRoot,
    
    [Parameter(Mandatory = $false)]
    [datetime]$StartDate,
    
    [Parameter(Mandatory = $false)]
    [datetime]$EndDate,
    
    [Parameter(Mandatory = $false)]
    [int]$TopN = 20,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeAllRequests
)

$ErrorActionPreference = 'Continue'

#region Configuration

$Script:ExchangeEndpointPatterns = @(
    'mapi',
    'microsoft-server-activesync',
    'ews',
    'owa',
    'autodiscover',
    'powershell',
    'oab',
    'ecp',
    'rpc'
)

# Regex pattern for faster matching (case-insensitive)
$Script:EndpointRegex = [regex]::new(
    '/(' + ($Script:ExchangeEndpointPatterns -join '|') + ')(/|$)',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Compiled
)

$Script:Thresholds = @{
    ResponseTime = @{
        Warning  = 500    # ms
        Critical = 2000   # ms
    }
    ErrorRate = @{
        Warning  = 2      # %
        Critical = 5      # %
    }
}

$Script:AnalysisResults = @{
    Timestamp       = Get-Date
    LogPath         = $LogPath
    TotalRequests   = 0
    ParsedFiles     = 0
    DateRange       = @{ Start = $null; End = $null }
    EndpointStats   = @{}
    HourlyVolume    = @{}
    SlowRequests    = @()
    UserActivity    = @{}
    StatusCodes     = @{}
    Issues          = @()
}

#endregion

#region Helper Functions

function Write-AnalyzerLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Info'    { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-StatusCategory {
    param([int]$StatusCode)
    
    switch ([math]::Floor($StatusCode / 100)) {
        2 { return 'Success' }
        3 { return 'Redirect' }
        4 { return 'ClientError' }
        5 { return 'ServerError' }
        default { return 'Unknown' }
    }
}

function Get-ThresholdColor {
    param(
        [double]$Value,
        [double]$WarningThreshold,
        [double]$CriticalThreshold
    )
    
    if ($Value -ge $CriticalThreshold) { return 'Critical' }
    elseif ($Value -ge $WarningThreshold) { return 'Warning' }
    else { return 'Healthy' }
}

function Get-Percentile {
    param(
        [double[]]$Values,
        [int]$Percentile
    )
    
    if ($Values.Count -eq 0) { return 0 }
    
    $sorted = $Values | Sort-Object
    $index = [math]::Ceiling(($Percentile / 100) * $sorted.Count) - 1
    $index = [math]::Max(0, [math]::Min($index, $sorted.Count - 1))
    
    return $sorted[$index]
}

function Get-EndpointFromUri {
    param([string]$Uri)
    
    $match = $Script:EndpointRegex.Match($Uri)
    if ($match.Success) {
        return $match.Groups[1].Value.ToUpper()
    }
    
    return 'Other'
}

function Test-IsExchangeEndpoint {
    param([string]$Uri)
    return $Script:EndpointRegex.IsMatch($Uri)
}

#endregion

#region IIS Log Parsing

function Get-IISLogFiles {
    param(
        [string]$Path,
        [datetime]$StartDate,
        [datetime]$EndDate
    )
    
    Write-AnalyzerLog "Scanning for IIS log files in: $Path" -Level Info
    
    $logFiles = @()
    
    if (Test-Path $Path -PathType Container) {
        $logFiles = Get-ChildItem -Path $Path -Filter "*.log" -Recurse -File | 
            Where-Object { $_.Length -gt 0 }
    }
    elseif (Test-Path $Path -PathType Leaf) {
        $logFiles = Get-Item $Path
    }
    else {
        Write-AnalyzerLog "Path not found: $Path" -Level Error
        return @()
    }
    
    if ($StartDate -or $EndDate) {
        $logFiles = $logFiles | Where-Object {
            $fileDate = $_.LastWriteTime.Date
            $include = $true
            if ($StartDate) { $include = $include -and ($fileDate -ge $StartDate.Date) }
            if ($EndDate) { $include = $include -and ($fileDate -le $EndDate.Date) }
            $include
        }
    }
    
    Write-AnalyzerLog "Found $($logFiles.Count) log file(s) to analyze" -Level Info
    return $logFiles
}

function Parse-IISLogFile {
    param(
        [System.IO.FileInfo]$LogFile
    )
    
    # Use ArrayList for much better performance than array += 
    $entries = [System.Collections.ArrayList]::new()
    $fieldNames = @()
    $fieldIndexes = @{}
    $lineCount = 0
    $matchCount = 0
    $lastProgressTime = [datetime]::Now
    
    try {
        $reader = [System.IO.StreamReader]::new($LogFile.FullName, [System.Text.Encoding]::UTF8, $true, 65536)
        
        while ($null -ne ($line = $reader.ReadLine())) {
            $lineCount++
            
            # Progress update every 2 seconds
            if (([datetime]::Now - $lastProgressTime).TotalSeconds -ge 2) {
                $lastProgressTime = [datetime]::Now
                Write-Host "`r  Processing: $($lineCount.ToString('N0')) lines, found $($matchCount.ToString('N0')) Exchange requests..." -NoNewline -ForegroundColor Gray
            }
            
            # Parse #Fields header
            if ($line.StartsWith('#Fields:')) {
                $fieldNames = $line.Substring(9).Trim() -split '\s+'
                # Build index lookup for faster field access
                $fieldIndexes = @{}
                for ($i = 0; $i -lt $fieldNames.Count; $i++) {
                    $fieldIndexes[$fieldNames[$i]] = $i
                }
                continue
            }
            
            # Skip comments and empty lines
            if ($line.StartsWith('#') -or $line.Length -eq 0) {
                continue
            }
            
            if ($fieldNames.Count -eq 0) {
                continue
            }
            
            $values = $line -split '\s+'
            
            if ($values.Count -lt $fieldNames.Count) {
                continue
            }
            
            # Get URI stem using index lookup (faster than hashtable per-line)
            $uriIndex = $fieldIndexes['cs-uri-stem']
            $uri = if ($null -ne $uriIndex -and $uriIndex -lt $values.Count) { $values[$uriIndex] } else { '' }
            
            # Fast regex check for Exchange endpoints
            if (-not $IncludeAllRequests) {
                if (-not (Test-IsExchangeEndpoint -Uri $uri)) { continue }
            }
            
            $matchCount++
            
            # Get field values using index lookup
            $dateVal = if ($fieldIndexes.ContainsKey('date')) { $values[$fieldIndexes['date']] } else { $null }
            $timeVal = if ($fieldIndexes.ContainsKey('time')) { $values[$fieldIndexes['time']] } else { $null }
            
            $parsedDateTime = $null
            if ($dateVal -and $timeVal) {
                try {
                    $parsedDateTime = [datetime]::ParseExact("$dateVal $timeVal", 'yyyy-MM-dd HH:mm:ss', $null)
                } catch { }
            }
            
            # Date filtering
            if ($StartDate -and $parsedDateTime -and $parsedDateTime -lt $StartDate) { continue }
            if ($EndDate -and $parsedDateTime -and $parsedDateTime -gt $EndDate.AddDays(1)) { continue }
            
            $parsedEntry = [PSCustomObject]@{
                DateTime    = $parsedDateTime
                ClientIP    = if ($fieldIndexes.ContainsKey('c-ip')) { $values[$fieldIndexes['c-ip']] } else { '-' }
                Username    = if ($fieldIndexes.ContainsKey('cs-username')) { $values[$fieldIndexes['cs-username']] } else { '-' }
                Method      = if ($fieldIndexes.ContainsKey('cs-method')) { $values[$fieldIndexes['cs-method']] } else { '-' }
                UriStem     = $uri
                UriQuery    = if ($fieldIndexes.ContainsKey('cs-uri-query')) { $values[$fieldIndexes['cs-uri-query']] } else { '-' }
                Status      = if ($fieldIndexes.ContainsKey('sc-status')) { [int]$values[$fieldIndexes['sc-status']] } else { 0 }
                SubStatus   = if ($fieldIndexes.ContainsKey('sc-substatus')) { [int]$values[$fieldIndexes['sc-substatus']] } else { 0 }
                TimeTaken   = if ($fieldIndexes.ContainsKey('time-taken')) { [int]$values[$fieldIndexes['time-taken']] } else { 0 }
                BytesSent   = if ($fieldIndexes.ContainsKey('sc-bytes')) { [long]$values[$fieldIndexes['sc-bytes']] } else { 0 }
                BytesRecv   = if ($fieldIndexes.ContainsKey('cs-bytes')) { [long]$values[$fieldIndexes['cs-bytes']] } else { 0 }
                UserAgent   = if ($fieldIndexes.ContainsKey('cs(User-Agent)')) { $values[$fieldIndexes['cs(User-Agent)']] } else { '-' }
                Endpoint    = Get-EndpointFromUri -Uri $uri
            }
            
            [void]$entries.Add($parsedEntry)
        }
        
        Write-Host "`r  Completed: $($lineCount.ToString('N0')) lines processed                    " -ForegroundColor Gray
        
        $reader.Close()
        $reader.Dispose()
        
    }
    catch {
        Write-Host ""
        Write-AnalyzerLog "Error parsing $($LogFile.Name): $($_.Exception.Message)" -Level Error
    }
    
    return $entries.ToArray()
}

#endregion

#region Analysis Functions

function Invoke-LogAnalysis {
    param(
        [array]$LogEntries
    )
    
    Write-AnalyzerLog "Analyzing $($LogEntries.Count) log entries..." -Level Info
    
    $Script:AnalysisResults.TotalRequests = $LogEntries.Count
    
    if ($LogEntries.Count -eq 0) {
        Write-AnalyzerLog "No log entries to analyze" -Level Warning
        return
    }
    
    $dates = $LogEntries | Where-Object { $_.DateTime } | Select-Object -ExpandProperty DateTime
    if ($dates) {
        $Script:AnalysisResults.DateRange.Start = ($dates | Measure-Object -Minimum).Minimum
        $Script:AnalysisResults.DateRange.End = ($dates | Measure-Object -Maximum).Maximum
    }
    
    $groupedByEndpoint = $LogEntries | Group-Object -Property Endpoint
    
    foreach ($group in $groupedByEndpoint) {
        $endpoint = $group.Name
        $requests = $group.Group
        $timeTakenValues = $requests | Select-Object -ExpandProperty TimeTaken
        
        $errorRequests = $requests | Where-Object { $_.Status -ge 400 }
        $errorRate = if ($requests.Count -gt 0) { 
            [math]::Round(($errorRequests.Count / $requests.Count) * 100, 2) 
        } else { 0 }
        
        $Script:AnalysisResults.EndpointStats[$endpoint] = @{
            TotalRequests = $requests.Count
            AvgResponseTime = [math]::Round(($timeTakenValues | Measure-Object -Average).Average, 2)
            MinResponseTime = ($timeTakenValues | Measure-Object -Minimum).Minimum
            MaxResponseTime = ($timeTakenValues | Measure-Object -Maximum).Maximum
            P50ResponseTime = Get-Percentile -Values $timeTakenValues -Percentile 50
            P95ResponseTime = Get-Percentile -Values $timeTakenValues -Percentile 95
            P99ResponseTime = Get-Percentile -Values $timeTakenValues -Percentile 99
            ErrorCount      = $errorRequests.Count
            ErrorRate       = $errorRate
            StatusCodes     = $requests | Group-Object -Property Status | 
                ForEach-Object { @{ $_.Name = $_.Count } }
        }
    }
    
    $LogEntries | Where-Object { $_.DateTime } | ForEach-Object {
        $hourKey = $_.DateTime.ToString('yyyy-MM-dd HH:00')
        $endpoint = $_.Endpoint
        
        if (-not $Script:AnalysisResults.HourlyVolume.ContainsKey($hourKey)) {
            $Script:AnalysisResults.HourlyVolume[$hourKey] = @{}
        }
        if (-not $Script:AnalysisResults.HourlyVolume[$hourKey].ContainsKey($endpoint)) {
            $Script:AnalysisResults.HourlyVolume[$hourKey][$endpoint] = 0
        }
        $Script:AnalysisResults.HourlyVolume[$hourKey][$endpoint]++
    }
    
    $Script:AnalysisResults.SlowRequests = $LogEntries | 
        Sort-Object -Property TimeTaken -Descending | 
        Select-Object -First $TopN
    
    $groupedByUser = $LogEntries | Where-Object { $_.Username -and $_.Username -ne '-' } | 
        Group-Object -Property Username
    
    foreach ($userGroup in $groupedByUser) {
        $userRequests = $userGroup.Group
        $userTimeTaken = $userRequests | Select-Object -ExpandProperty TimeTaken
        
        $Script:AnalysisResults.UserActivity[$userGroup.Name] = @{
            TotalRequests   = $userRequests.Count
            AvgResponseTime = [math]::Round(($userTimeTaken | Measure-Object -Average).Average, 2)
            Endpoints       = ($userRequests | Group-Object -Property Endpoint | 
                ForEach-Object { "$($_.Name):$($_.Count)" }) -join ', '
        }
    }
    
    $statusGroups = $LogEntries | Group-Object -Property Status
    foreach ($sg in $statusGroups) {
        $Script:AnalysisResults.StatusCodes[$sg.Name] = $sg.Count
    }
    
    foreach ($endpoint in $Script:AnalysisResults.EndpointStats.Keys) {
        $stats = $Script:AnalysisResults.EndpointStats[$endpoint]
        
        if ($stats.AvgResponseTime -ge $Script:Thresholds.ResponseTime.Critical) {
            $Script:AnalysisResults.Issues += @{
                Level    = 'Critical'
                Endpoint = $endpoint
                Message  = "Average response time ($($stats.AvgResponseTime)ms) exceeds critical threshold ($($Script:Thresholds.ResponseTime.Critical)ms)"
            }
        }
        elseif ($stats.AvgResponseTime -ge $Script:Thresholds.ResponseTime.Warning) {
            $Script:AnalysisResults.Issues += @{
                Level    = 'Warning'
                Endpoint = $endpoint
                Message  = "Average response time ($($stats.AvgResponseTime)ms) exceeds warning threshold ($($Script:Thresholds.ResponseTime.Warning)ms)"
            }
        }
        
        if ($stats.ErrorRate -ge $Script:Thresholds.ErrorRate.Critical) {
            $Script:AnalysisResults.Issues += @{
                Level    = 'Critical'
                Endpoint = $endpoint
                Message  = "Error rate ($($stats.ErrorRate)%) exceeds critical threshold ($($Script:Thresholds.ErrorRate.Critical)%)"
            }
        }
        elseif ($stats.ErrorRate -ge $Script:Thresholds.ErrorRate.Warning) {
            $Script:AnalysisResults.Issues += @{
                Level    = 'Warning'
                Endpoint = $endpoint
                Message  = "Error rate ($($stats.ErrorRate)%) exceeds warning threshold ($($Script:Thresholds.ErrorRate.Warning)%)"
            }
        }
    }
    
    Write-AnalyzerLog "Analysis complete" -Level Success
}

#endregion

#region Output Functions

function Write-ConsoleReport {
    Write-Host ""
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "  EXCHANGE IIS LOG ANALYSIS REPORT" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Summary" -ForegroundColor White
    Write-Host "-" * 40
    Write-Host "  Log Path:        $($Script:AnalysisResults.LogPath)"
    Write-Host "  Files Analyzed:  $($Script:AnalysisResults.ParsedFiles)"
    Write-Host "  Total Requests:  $($Script:AnalysisResults.TotalRequests)"
    if ($Script:AnalysisResults.DateRange.Start) {
        Write-Host "  Date Range:      $($Script:AnalysisResults.DateRange.Start.ToString('yyyy-MM-dd HH:mm')) to $($Script:AnalysisResults.DateRange.End.ToString('yyyy-MM-dd HH:mm'))"
    }
    Write-Host ""
    
    if ($Script:AnalysisResults.Issues.Count -gt 0) {
        Write-Host "Issues Detected" -ForegroundColor Red
        Write-Host "-" * 40
        foreach ($issue in $Script:AnalysisResults.Issues) {
            $color = if ($issue.Level -eq 'Critical') { 'Red' } else { 'Yellow' }
            Write-Host "  [$($issue.Level)] $($issue.Endpoint): $($issue.Message)" -ForegroundColor $color
        }
        Write-Host ""
    }
    
    Write-Host "Endpoint Performance" -ForegroundColor White
    Write-Host "-" * 40
    
    $tableData = @()
    foreach ($endpoint in $Script:AnalysisResults.EndpointStats.Keys | Sort-Object) {
        $stats = $Script:AnalysisResults.EndpointStats[$endpoint]
        $tableData += [PSCustomObject]@{
            Endpoint    = $endpoint
            Requests    = $stats.TotalRequests
            'Avg(ms)'   = $stats.AvgResponseTime
            'P95(ms)'   = $stats.P95ResponseTime
            'P99(ms)'   = $stats.P99ResponseTime
            'Max(ms)'   = $stats.MaxResponseTime
            'Errors'    = $stats.ErrorCount
            'Err%'      = "$($stats.ErrorRate)%"
        }
    }
    
    $tableData | Format-Table -AutoSize | Out-String | Write-Host
    
    Write-Host "Top $TopN Slowest Requests" -ForegroundColor White
    Write-Host "-" * 40
    
    $slowTable = $Script:AnalysisResults.SlowRequests | Select-Object @{N='Time';E={$_.DateTime.ToString('MM-dd HH:mm:ss')}}, 
        @{N='ms';E={$_.TimeTaken}}, 
        @{N='Status';E={$_.Status}},
        @{N='Endpoint';E={$_.Endpoint}},
        @{N='User';E={if($_.Username -eq '-'){'(anonymous)'}else{$_.Username}}},
        @{N='URI';E={$_.UriStem.Substring(0, [Math]::Min(50, $_.UriStem.Length))}}
    
    $slowTable | Format-Table -AutoSize | Out-String | Write-Host
    
    Write-Host "HTTP Status Code Distribution" -ForegroundColor White
    Write-Host "-" * 40
    
    $statusTable = $Script:AnalysisResults.StatusCodes.GetEnumerator() | 
        Sort-Object -Property Name | 
        ForEach-Object {
            $category = Get-StatusCategory -StatusCode ([int]$_.Name)
            [PSCustomObject]@{
                Status   = $_.Name
                Count    = $_.Value
                Category = $category
                Percent  = [math]::Round(($_.Value / $Script:AnalysisResults.TotalRequests) * 100, 2)
            }
        }
    
    $statusTable | Format-Table -AutoSize | Out-String | Write-Host
    
    if ($Script:AnalysisResults.UserActivity.Count -gt 0) {
        Write-Host "Top 10 Users by Request Volume" -ForegroundColor White
        Write-Host "-" * 40
        
        $userTable = $Script:AnalysisResults.UserActivity.GetEnumerator() | 
            Sort-Object { $_.Value.TotalRequests } -Descending | 
            Select-Object -First 10 | 
            ForEach-Object {
                [PSCustomObject]@{
                    User        = $_.Name
                    Requests    = $_.Value.TotalRequests
                    'Avg(ms)'   = $_.Value.AvgResponseTime
                }
            }
        
        $userTable | Format-Table -AutoSize | Out-String | Write-Host
    }
}

function Export-HTMLReport {
    param([string]$OutputFile)
    
    $criticalColor = '#dc3545'
    $warningColor = '#ffc107'
    $healthyColor = '#28a745'
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange IIS Log Analysis Report</title>
    <style>
        * { box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: #f5f5f5; 
            color: #333;
        }
        .container { max-width: 1400px; margin: 0 auto; }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #444; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
        .summary-box { 
            background: white; 
            padding: 20px; 
            border-radius: 8px; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .summary-item { text-align: center; padding: 15px; background: #f8f9fa; border-radius: 5px; }
        .summary-item .value { font-size: 2em; font-weight: bold; color: #0078d4; }
        .summary-item .label { color: #666; font-size: 0.9em; }
        .issue { padding: 10px 15px; margin: 5px 0; border-radius: 5px; }
        .issue.critical { background: #f8d7da; border-left: 4px solid $criticalColor; }
        .issue.warning { background: #fff3cd; border-left: 4px solid $warningColor; }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            background: white; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background: #0078d4; color: white; font-weight: 600; }
        tr:hover { background: #f5f5f5; }
        tr:nth-child(even) { background: #fafafa; }
        .status-healthy { color: $healthyColor; font-weight: bold; }
        .status-warning { color: $warningColor; font-weight: bold; }
        .status-critical { color: $criticalColor; font-weight: bold; }
        .timestamp { color: #666; font-size: 0.9em; margin-bottom: 20px; }
        .footer { text-align: center; color: #666; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Exchange IIS Log Analysis Report</h1>
        <p class="timestamp">Generated: $($Script:AnalysisResults.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        
        <div class="summary-box">
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="value">$($Script:AnalysisResults.TotalRequests.ToString('N0'))</div>
                    <div class="label">Total Requests</div>
                </div>
                <div class="summary-item">
                    <div class="value">$($Script:AnalysisResults.ParsedFiles)</div>
                    <div class="label">Files Analyzed</div>
                </div>
                <div class="summary-item">
                    <div class="value">$($Script:AnalysisResults.EndpointStats.Count)</div>
                    <div class="label">Endpoints</div>
                </div>
                <div class="summary-item">
                    <div class="value">$(($Script:AnalysisResults.Issues | Where-Object { $_.Level -eq 'Critical' }).Count)</div>
                    <div class="label">Critical Issues</div>
                </div>
            </div>
        </div>
"@

    if ($Script:AnalysisResults.DateRange.Start) {
        $html += @"
        <p><strong>Date Range:</strong> $($Script:AnalysisResults.DateRange.Start.ToString('yyyy-MM-dd HH:mm')) to $($Script:AnalysisResults.DateRange.End.ToString('yyyy-MM-dd HH:mm'))</p>
        <p><strong>Log Path:</strong> $($Script:AnalysisResults.LogPath)</p>
"@
    }

    if ($Script:AnalysisResults.Issues.Count -gt 0) {
        $html += "<h2>Issues Detected</h2>"
        foreach ($issue in $Script:AnalysisResults.Issues) {
            $cssClass = $issue.Level.ToLower()
            $html += "<div class='issue $cssClass'><strong>[$($issue.Level)] $($issue.Endpoint):</strong> $($issue.Message)</div>"
        }
    }

    $html += @"
        <h2>Endpoint Performance</h2>
        <table>
            <tr>
                <th>Endpoint</th>
                <th>Requests</th>
                <th>Avg (ms)</th>
                <th>P50 (ms)</th>
                <th>P95 (ms)</th>
                <th>P99 (ms)</th>
                <th>Max (ms)</th>
                <th>Errors</th>
                <th>Error %</th>
                <th>Status</th>
            </tr>
"@

    foreach ($endpoint in $Script:AnalysisResults.EndpointStats.Keys | Sort-Object) {
        $stats = $Script:AnalysisResults.EndpointStats[$endpoint]
        $status = Get-ThresholdColor -Value $stats.AvgResponseTime -WarningThreshold $Script:Thresholds.ResponseTime.Warning -CriticalThreshold $Script:Thresholds.ResponseTime.Critical
        $statusClass = "status-$($status.ToLower())"
        
        $html += @"
            <tr>
                <td><strong>$endpoint</strong></td>
                <td>$($stats.TotalRequests.ToString('N0'))</td>
                <td>$($stats.AvgResponseTime)</td>
                <td>$($stats.P50ResponseTime)</td>
                <td>$($stats.P95ResponseTime)</td>
                <td>$($stats.P99ResponseTime)</td>
                <td>$($stats.MaxResponseTime)</td>
                <td>$($stats.ErrorCount)</td>
                <td>$($stats.ErrorRate)%</td>
                <td class="$statusClass">$status</td>
            </tr>
"@
    }

    $html += "</table>"

    $html += @"
        <h2>Top $TopN Slowest Requests</h2>
        <table>
            <tr>
                <th>Time</th>
                <th>Duration (ms)</th>
                <th>Status</th>
                <th>Endpoint</th>
                <th>User</th>
                <th>URI</th>
            </tr>
"@

    foreach ($req in $Script:AnalysisResults.SlowRequests) {
        $timeStr = if ($req.DateTime) { $req.DateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'N/A' }
        $user = if ($req.Username -eq '-') { '(anonymous)' } else { $req.Username }
        $uriTrunc = if ($req.UriStem.Length -gt 60) { $req.UriStem.Substring(0, 60) + '...' } else { $req.UriStem }
        
        $html += @"
            <tr>
                <td>$timeStr</td>
                <td><strong>$($req.TimeTaken)</strong></td>
                <td>$($req.Status)</td>
                <td>$($req.Endpoint)</td>
                <td>$user</td>
                <td title="$($req.UriStem)">$uriTrunc</td>
            </tr>
"@
    }

    $html += "</table>"

    $html += @"
        <h2>HTTP Status Code Distribution</h2>
        <table>
            <tr>
                <th>Status Code</th>
                <th>Count</th>
                <th>Percentage</th>
                <th>Category</th>
            </tr>
"@

    foreach ($status in $Script:AnalysisResults.StatusCodes.GetEnumerator() | Sort-Object Name) {
        $category = Get-StatusCategory -StatusCode ([int]$status.Name)
        $percent = [math]::Round(($status.Value / $Script:AnalysisResults.TotalRequests) * 100, 2)
        $categoryClass = switch ($category) {
            'Success' { 'status-healthy' }
            'ClientError' { 'status-warning' }
            'ServerError' { 'status-critical' }
            default { '' }
        }
        
        $html += @"
            <tr>
                <td><strong>$($status.Name)</strong></td>
                <td>$($status.Value.ToString('N0'))</td>
                <td>$percent%</td>
                <td class="$categoryClass">$category</td>
            </tr>
"@
    }

    $html += "</table>"

    if ($Script:AnalysisResults.UserActivity.Count -gt 0) {
        $html += @"
        <h2>Top Users by Request Volume</h2>
        <table>
            <tr>
                <th>User</th>
                <th>Requests</th>
                <th>Avg Response (ms)</th>
                <th>Endpoints Used</th>
            </tr>
"@

        $topUsers = $Script:AnalysisResults.UserActivity.GetEnumerator() | 
            Sort-Object { $_.Value.TotalRequests } -Descending | 
            Select-Object -First 20

        foreach ($user in $topUsers) {
            $html += @"
            <tr>
                <td><strong>$($user.Name)</strong></td>
                <td>$($user.Value.TotalRequests.ToString('N0'))</td>
                <td>$($user.Value.AvgResponseTime)</td>
                <td>$($user.Value.Endpoints)</td>
            </tr>
"@
        }

        $html += "</table>"
    }

    $html += @"
        <div class="footer">
            <p>Exchange IIS Log Analyzer | Report generated on $($Script:AnalysisResults.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-AnalyzerLog "HTML report saved to: $OutputFile" -Level Success
}

function Export-CSVReport {
    param([string]$OutputPath)
    
    $endpointCsv = Join-Path $OutputPath "ExchangeIIS_EndpointStats_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $slowRequestsCsv = Join-Path $OutputPath "ExchangeIIS_SlowRequests_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $statusCodesCsv = Join-Path $OutputPath "ExchangeIIS_StatusCodes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $userActivityCsv = Join-Path $OutputPath "ExchangeIIS_UserActivity_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    $endpointData = foreach ($endpoint in $Script:AnalysisResults.EndpointStats.Keys) {
        $stats = $Script:AnalysisResults.EndpointStats[$endpoint]
        [PSCustomObject]@{
            Endpoint        = $endpoint
            TotalRequests   = $stats.TotalRequests
            AvgResponseMs   = $stats.AvgResponseTime
            MinResponseMs   = $stats.MinResponseTime
            MaxResponseMs   = $stats.MaxResponseTime
            P50ResponseMs   = $stats.P50ResponseTime
            P95ResponseMs   = $stats.P95ResponseTime
            P99ResponseMs   = $stats.P99ResponseTime
            ErrorCount      = $stats.ErrorCount
            ErrorRate       = $stats.ErrorRate
        }
    }
    $endpointData | Export-Csv -Path $endpointCsv -NoTypeInformation
    Write-AnalyzerLog "Endpoint stats exported to: $endpointCsv" -Level Success
    
    $Script:AnalysisResults.SlowRequests | Select-Object DateTime, TimeTaken, Status, SubStatus, 
        Endpoint, UriStem, Username, ClientIP, Method, BytesSent, BytesRecv |
        Export-Csv -Path $slowRequestsCsv -NoTypeInformation
    Write-AnalyzerLog "Slow requests exported to: $slowRequestsCsv" -Level Success
    
    $statusData = foreach ($status in $Script:AnalysisResults.StatusCodes.GetEnumerator()) {
        [PSCustomObject]@{
            StatusCode = $status.Name
            Count      = $status.Value
            Percentage = [math]::Round(($status.Value / $Script:AnalysisResults.TotalRequests) * 100, 2)
            Category   = Get-StatusCategory -StatusCode ([int]$status.Name)
        }
    }
    $statusData | Export-Csv -Path $statusCodesCsv -NoTypeInformation
    Write-AnalyzerLog "Status codes exported to: $statusCodesCsv" -Level Success
    
    if ($Script:AnalysisResults.UserActivity.Count -gt 0) {
        $userData = foreach ($user in $Script:AnalysisResults.UserActivity.GetEnumerator()) {
            [PSCustomObject]@{
                Username        = $user.Name
                TotalRequests   = $user.Value.TotalRequests
                AvgResponseMs   = $user.Value.AvgResponseTime
                Endpoints       = $user.Value.Endpoints
            }
        }
        $userData | Export-Csv -Path $userActivityCsv -NoTypeInformation
        Write-AnalyzerLog "User activity exported to: $userActivityCsv" -Level Success
    }
}

#endregion

#region Main Execution

Write-Host ""
Write-AnalyzerLog "Exchange IIS Log Analyzer v1.0" -Level Info
Write-AnalyzerLog "Starting analysis..." -Level Info

$logFiles = Get-IISLogFiles -Path $LogPath -StartDate $StartDate -EndDate $EndDate

if ($logFiles.Count -eq 0) {
    Write-AnalyzerLog "No log files found matching criteria. Exiting." -Level Error
    exit 1
}

$allEntries = [System.Collections.ArrayList]::new()

foreach ($logFile in $logFiles) {
    Write-AnalyzerLog "Parsing: $($logFile.Name) ($([math]::Round($logFile.Length / 1MB, 2)) MB)" -Level Info
    $entries = Parse-IISLogFile -LogFile $logFile
    if ($entries.Count -gt 0) {
        [void]$allEntries.AddRange($entries)
    }
    $Script:AnalysisResults.ParsedFiles++
    Write-AnalyzerLog "  Found $($entries.Count) Exchange-related entries" -Level Info
}

Invoke-LogAnalysis -LogEntries $allEntries

Write-ConsoleReport

if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$htmlFile = Join-Path $OutputPath "ExchangeIIS_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
Export-HTMLReport -OutputFile $htmlFile

Export-CSVReport -OutputPath $OutputPath

Write-Host ""
Write-AnalyzerLog "Analysis complete!" -Level Success
Write-Host ""

#endregion
