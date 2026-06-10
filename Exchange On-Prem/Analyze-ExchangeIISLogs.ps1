#Requires -Version 5.1
<#
.SYNOPSIS
    Analyzes IIS and Exchange HttpProxy logs from Exchange 2016 servers to identify
    performance issues, errors, and top talkers.

.DESCRIPTION
    This script parses IIS W3C logs and optionally Exchange HttpProxy logs from one or
    more Exchange 2016 servers. It identifies slow requests, HTTP errors, authentication
    failures, and the users/clients generating the most load. Results are presented in
    a color-coded HTML report and optional CSV export.

    All Exchange 2016 client traffic (Outlook/MAPI, OWA, EWS, ActiveSync, Autodiscover)
    flows through IIS, making these logs essential for diagnosing client connectivity
    and performance issues.

.PARAMETER Servers
    Mandatory. One or more Exchange server names to analyze.

.PARAMETER Hours
    Number of hours back to analyze. Default: 24.

.PARAMETER OutputPath
    Directory for report output. Default: script directory.

.PARAMETER SlowRequestThreshold
    Requests taking longer than this (milliseconds) are flagged as slow. Default: 5000.

.PARAMETER LogSite
    Which IIS site logs to analyze: Frontend (W3SVC1), Backend (W3SVC2), or Both. Default: Both.

.PARAMETER IncludeHttpProxy
    Switch to also analyze Exchange HttpProxy logs for deeper backend latency analysis.

.PARAMETER RunLocally
    Switch to use local file paths (C:\) instead of UNC paths (\\Server\C$).
    Use this when running the script directly on each Exchange server for much faster
    log parsing (avoids SMB overhead). When enabled, only the local server is analyzed
    regardless of how many servers are specified in -Servers.

.PARAMETER SkipIISLogs
    Switch to skip IIS log analysis (useful if you only want HttpProxy analysis).

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -Servers "EX01","EX02"
    Analyzes IIS logs from the last 24 hours on both servers.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -Servers "EX01" -Hours 4 -IncludeHttpProxy
    Analyzes last 4 hours of IIS and HttpProxy logs on EX01.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -Servers "EX01","EX02" -SlowRequestThreshold 2000
    Flags requests slower than 2 seconds.

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -Servers $env:COMPUTERNAME -RunLocally
    Runs locally on the current Exchange server using local paths (fastest).

.EXAMPLE
    .\Analyze-ExchangeIISLogs.ps1 -Servers $env:COMPUTERNAME -RunLocally -IncludeHttpProxy -Hours 4
    Local analysis of last 4 hours including HttpProxy logs.

.NOTES
    Author: Exchange Performance Diagnostics Toolkit
    Requires: Network access to server admin shares (C$)
    No external dependencies (no LogParser required)
#>

param(
    [Parameter(Mandatory = $true)]
    [string[]]$Servers,

    [int]$Hours = 24,

    [string]$OutputPath = $PSScriptRoot,

    [int]$SlowRequestThreshold = 5000,

    [ValidateSet('Frontend', 'Backend', 'Both')]
    [string]$LogSite = 'Both',

    [switch]$IncludeHttpProxy,

    [switch]$SkipIISLogs,

    [switch]$RunLocally
)

#region Helper Functions

function Write-AnalysisLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'HH:mm:ss'
    $colors = @{
        'Info'    = 'Cyan'
        'Warning' = 'Yellow'
        'Error'   = 'Red'
        'Success' = 'Green'
    }
    
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
    Write-Host $Message -ForegroundColor $colors[$Level]
}

function Get-IISLogFields {
    <#
    .SYNOPSIS
        Parses the #Fields header line from a W3C IIS log file and returns field names.
    #>
    param([string]$FieldsLine)
    
    if ($FieldsLine -match '^#Fields:\s+(.+)$') {
        return ($Matches[1] -split '\s+')
    }
    return $null
}

function ConvertFrom-IISLogLine {
    <#
    .SYNOPSIS
        Parses a single IIS log line into a hashtable using the field definitions.
    #>
    param(
        [string]$Line,
        [string[]]$Fields
    )
    
    if ($Line.StartsWith('#')) { return $null }
    
    $values = $Line -split '\s+'
    if ($values.Count -ne $Fields.Count) { return $null }
    
    $entry = @{}
    for ($i = 0; $i -lt $Fields.Count; $i++) {
        $entry[$Fields[$i]] = $values[$i]
    }
    
    return $entry
}

function Get-VirtualDirectoryName {
    <#
    .SYNOPSIS
        Extracts the Exchange virtual directory name from a URI stem.
    #>
    param([string]$UriStem)
    
    $knownVDirs = @(
        '/mapi/', '/rpc/', '/owa/', '/ews/',
        '/Microsoft-Server-ActiveSync/', '/autodiscover/',
        '/oab/', '/powershell/', '/ecp/'
    )
    
    foreach ($vdir in $knownVDirs) {
        if ($UriStem -like "$vdir*" -or $UriStem -eq $vdir.TrimEnd('/')) {
            return $vdir.Trim('/')
        }
    }
    
    return 'Other'
}

function Get-FriendlyVDirName {
    <#
    .SYNOPSIS
        Returns a human-friendly name for an Exchange virtual directory.
    #>
    param([string]$VDir)
    
    $map = @{
        'mapi'                          = 'MAPI/HTTP (Outlook)'
        'rpc'                           = 'RPC/HTTP (Outlook Anywhere)'
        'owa'                           = 'Outlook Web App'
        'ews'                           = 'Exchange Web Services'
        'Microsoft-Server-ActiveSync'   = 'ActiveSync (Mobile)'
        'autodiscover'                  = 'Autodiscover'
        'oab'                           = 'Offline Address Book'
        'powershell'                    = 'Remote PowerShell'
        'ecp'                           = 'Exchange Admin Center'
        'Other'                         = 'Other'
    }
    
    if ($map.ContainsKey($VDir)) { return $map[$VDir] }
    return $VDir
}

function Get-PercentileValue {
    <#
    .SYNOPSIS
        Calculates the Nth percentile from a sorted array of numbers.
    #>
    param(
        [double[]]$SortedValues,
        [int]$Percentile = 95
    )
    
    if ($SortedValues.Count -eq 0) { return 0 }
    if ($SortedValues.Count -eq 1) { return $SortedValues[0] }
    
    $index = [math]::Ceiling(($Percentile / 100.0) * $SortedValues.Count) - 1
    $index = [math]::Max(0, [math]::Min($index, $SortedValues.Count - 1))
    return $SortedValues[$index]
}

#endregion

#region IIS Log Parsing

function Read-IISLogFiles {
    <#
    .SYNOPSIS
        Reads and parses IIS log files from a server for the specified time range.
        Returns an array of parsed log entries.
    #>
    param(
        [string]$ServerName,
        [string]$SitePath,
        [datetime]$StartTime,
        [int]$SlowThreshold
    )
    
    $results = @{
        Entries       = [System.Collections.Generic.List[PSObject]]::new()
        TotalLines    = 0
        ParseErrors   = 0
        FilesProcessed = 0
        Status        = 'Unknown'
    }
    
    $logPath = "\\$ServerName\C`$\inetpub\logs\LogFiles\$SitePath"
    
    if (-not (Test-Path -Path $logPath -PathType Container)) {
        $results.Status = 'PathNotFound'
        return $results
    }
    
    # Get log files modified within the time range (plus buffer)
    $logFiles = Get-ChildItem -Path $logPath -Filter '*.log' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $StartTime.AddHours(-1) } |
        Sort-Object LastWriteTime
    
    if (-not $logFiles) {
        $results.Status = 'NoFilesFound'
        return $results
    }
    
    $startDateStr = $StartTime.ToString('yyyy-MM-dd')
    $startTimeStr = $StartTime.ToString('HH:mm:ss')
    
    foreach ($logFile in $logFiles) {
        $results.FilesProcessed++
        $fields = $null
        
        try {
            $reader = [System.IO.StreamReader]::new($logFile.FullName)
            
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                $results.TotalLines++
                
                # Parse field definitions
                if ($line.StartsWith('#Fields:')) {
                    $fields = Get-IISLogFields -FieldsLine $line
                    continue
                }
                
                # Skip other comment lines
                if ($line.StartsWith('#')) { continue }
                if ($null -eq $fields) { continue }
                
                # Quick date filter before full parsing
                $lineParts = $line -split '\s+', 3
                if ($lineParts.Count -lt 2) { continue }
                
                $lineDate = $lineParts[0]
                $lineTime = $lineParts[1]
                
                # Skip entries before our start time
                if ($lineDate -lt $startDateStr) { continue }
                if ($lineDate -eq $startDateStr -and $lineTime -lt $startTimeStr) { continue }
                
                # Full parse
                $entry = ConvertFrom-IISLogLine -Line $line -Fields $fields
                if ($null -eq $entry) {
                    $results.ParseErrors++
                    continue
                }
                
                # Build structured entry
                $timeTaken = 0
                if ($entry.ContainsKey('time-taken')) {
                    [int]::TryParse($entry['time-taken'], [ref]$timeTaken) | Out-Null
                }
                
                $statusCode = 0
                if ($entry.ContainsKey('sc-status')) {
                    [int]::TryParse($entry['sc-status'], [ref]$statusCode) | Out-Null
                }
                
                $uriStem = if ($entry.ContainsKey('cs-uri-stem')) { $entry['cs-uri-stem'] } else { '/' }
                $vdir = Get-VirtualDirectoryName -UriStem $uriStem
                
                $parsedEntry = [PSCustomObject]@{
                    DateTime    = "$($entry['date']) $($entry['time'])"
                    Method      = if ($entry.ContainsKey('cs-method')) { $entry['cs-method'] } else { '-' }
                    UriStem     = $uriStem
                    VDir        = $vdir
                    StatusCode  = $statusCode
                    SubStatus   = if ($entry.ContainsKey('sc-substatus')) { $entry['sc-substatus'] } else { '0' }
                    TimeTaken   = $timeTaken
                    ClientIP    = if ($entry.ContainsKey('c-ip')) { $entry['c-ip'] } else { '-' }
                    Username    = if ($entry.ContainsKey('cs-username')) { $entry['cs-username'] } else { '-' }
                    UserAgent   = if ($entry.ContainsKey('cs(User-Agent)')) { $entry['cs(User-Agent)'] } else { '-' }
                    ServerIP    = if ($entry.ContainsKey('s-ip')) { $entry['s-ip'] } else { '-' }
                    ServerPort  = if ($entry.ContainsKey('s-port')) { $entry['s-port'] } else { '-' }
                    BytesSent   = if ($entry.ContainsKey('sc-bytes')) { $entry['sc-bytes'] } else { '0' }
                    BytesRecv   = if ($entry.ContainsKey('cs-bytes')) { $entry['cs-bytes'] } else { '0' }
                    IsSlow      = ($timeTaken -ge $SlowThreshold)
                    IsError     = ($statusCode -ge 400)
                }
                
                $results.Entries.Add($parsedEntry)
            }
            
            $reader.Close()
            $reader.Dispose()
            
        } catch {
            if ($reader) { $reader.Dispose() }
            $results.ParseErrors++
        }
    }
    
    $results.Status = 'Collected'
    return $results
}

#endregion

#region HttpProxy Log Parsing

function Read-HttpProxyLogFiles {
    <#
    .SYNOPSIS
        Reads and parses Exchange HttpProxy log files from a server.
        HttpProxy logs contain detailed backend latency information.
    #>
    param(
        [string]$ServerName,
        [string]$ProxyType,
        [datetime]$StartTime
    )
    
    $results = @{
        Entries        = [System.Collections.Generic.List[PSObject]]::new()
        TotalLines     = 0
        ParseErrors    = 0
        FilesProcessed = 0
        Status         = 'Unknown'
    }
    
    $basePath = "\\$ServerName\C`$\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\$ProxyType"
    
    if (-not (Test-Path -Path $basePath -PathType Container)) {
        $results.Status = 'PathNotFound'
        return $results
    }
    
    $logFiles = Get-ChildItem -Path $basePath -Filter '*.log' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $StartTime.AddHours(-1) } |
        Sort-Object LastWriteTime
    
    if (-not $logFiles) {
        $results.Status = 'NoFilesFound'
        return $results
    }
    
    foreach ($logFile in $logFiles) {
        $results.FilesProcessed++
        $headers = $null
        
        try {
            $reader = [System.IO.StreamReader]::new($logFile.FullName)
            
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                $results.TotalLines++
                
                # First non-comment line with #DateTime is the header
                if ($line.StartsWith('#') -and $line -match '^#.*DateTime') {
                    $headers = ($line.TrimStart('#') -split ',')
                    continue
                }
                
                if ($line.StartsWith('#')) { continue }
                if ($null -eq $headers) { continue }
                
                $values = $line -split ','
                if ($values.Count -lt $headers.Count) {
                    $results.ParseErrors++
                    continue
                }
                
                $entry = @{}
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    $entry[$headers[$i].Trim()] = $values[$i].Trim()
                }
                
                # Date filter
                $dateTimeStr = if ($entry.ContainsKey('DateTime')) { $entry['DateTime'] } else { $null }
                if ($null -eq $dateTimeStr) { continue }
                
                $entryTime = [datetime]::MinValue
                if (-not [datetime]::TryParse($dateTimeStr, [ref]$entryTime)) { continue }
                if ($entryTime -lt $StartTime) { continue }
                
                # Parse latency values
                $totalLatency = 0
                $backendLatency = 0
                $authLatency = 0
                
                if ($entry.ContainsKey('TotalRequestTime')) {
                    [int]::TryParse($entry['TotalRequestTime'], [ref]$totalLatency) | Out-Null
                }
                if ($entry.ContainsKey('BackendProcessingLatency')) {
                    [int]::TryParse($entry['BackendProcessingLatency'], [ref]$backendLatency) | Out-Null
                }
                if ($entry.ContainsKey('AuthModuleLatency')) {
                    [int]::TryParse($entry['AuthModuleLatency'], [ref]$authLatency) | Out-Null
                }
                
                $statusCode = 0
                if ($entry.ContainsKey('HttpStatus')) {
                    [int]::TryParse($entry['HttpStatus'], [ref]$statusCode) | Out-Null
                }
                
                $parsedEntry = [PSCustomObject]@{
                    DateTime              = $entryTime
                    ProxyType             = $ProxyType
                    AuthenticatedUser     = if ($entry.ContainsKey('AuthenticatedUser')) { $entry['AuthenticatedUser'] } else { '-' }
                    UrlStem               = if ($entry.ContainsKey('UrlStem')) { $entry['UrlStem'] } else { '-' }
                    HttpStatus            = $statusCode
                    BackendServer         = if ($entry.ContainsKey('TargetServer')) { $entry['TargetServer'] } else { '-' }
                    TotalLatency          = $totalLatency
                    BackendLatency        = $backendLatency
                    AuthLatency           = $authLatency
                    ClientIP              = if ($entry.ContainsKey('ClientIpAddress')) { $entry['ClientIpAddress'] } else { '-' }
                    ErrorCode             = if ($entry.ContainsKey('ErrorCode')) { $entry['ErrorCode'] } else { '-' }
                    GenericErrors         = if ($entry.ContainsKey('GenericErrors')) { $entry['GenericErrors'] } else { '-' }
                    ServerLocatorLatency  = if ($entry.ContainsKey('ServerLocatorLatency')) { $entry['ServerLocatorLatency'] } else { '0' }
                    TargetServerVersion   = if ($entry.ContainsKey('TargetServerVersion')) { $entry['TargetServerVersion'] } else { '-' }
                    RequestId             = if ($entry.ContainsKey('RequestId')) { $entry['RequestId'] } else { '-' }
                }
                
                $results.Entries.Add($parsedEntry)
            }
            
            $reader.Close()
            $reader.Dispose()
            
        } catch {
            if ($reader) { $reader.Dispose() }
            $results.ParseErrors++
        }
    }
    
    $results.Status = 'Collected'
    return $results
}

#endregion

#region Analysis Functions

function Get-IISAnalysis {
    <#
    .SYNOPSIS
        Analyzes parsed IIS log entries and returns structured analysis results.
    #>
    param(
        [System.Collections.Generic.List[PSObject]]$Entries,
        [int]$SlowThreshold
    )
    
    $analysis = @{
        TotalRequests     = $Entries.Count
        SlowRequests      = 0
        ErrorRequests     = 0
        AvgResponseTime   = 0
        P95ResponseTime   = 0
        MaxResponseTime   = 0
        StatusCodes       = @{}
        VDirStats         = @{}
        TopUsersByCount   = @()
        TopUsersByLatency = @()
        TopClientIPs      = @()
        RequestsPerHour   = @{}
        ErrorsPerHour     = @{}
        SlowByVDir        = @{}
        AuthFailures      = @()
        ServerErrors      = @()
    }
    
    if ($Entries.Count -eq 0) { return $analysis }
    
    # Collect all time-taken values for percentile calculation
    $allTimeTaken = [System.Collections.Generic.List[double]]::new()
    
    # Per-user tracking
    $userRequests = @{}
    $userLatency = @{}
    $clientIPCount = @{}
    
    $entryCount = $Entries.Count
    $processedCount = 0
    
    foreach ($entry in $Entries) {
        $processedCount++
        if ($processedCount % 5000 -eq 0 -or $processedCount -eq $entryCount) {
            $pct = [math]::Round(($processedCount / $entryCount) * 100)
            Write-Progress -Id 2 -ParentId 0 -Activity "Analyzing IIS entries" `
                -Status "$processedCount of $entryCount entries ($pct%)" -PercentComplete $pct
        }
        $allTimeTaken.Add($entry.TimeTaken)
        
        # Slow request tracking
        if ($entry.IsSlow) { $analysis.SlowRequests++ }
        if ($entry.IsError) { $analysis.ErrorRequests++ }
        
        # Status code distribution
        $sc = $entry.StatusCode.ToString()
        if (-not $analysis.StatusCodes.ContainsKey($sc)) { $analysis.StatusCodes[$sc] = 0 }
        $analysis.StatusCodes[$sc]++
        
        # Virtual directory stats
        $vdir = $entry.VDir
        if (-not $analysis.VDirStats.ContainsKey($vdir)) {
            $analysis.VDirStats[$vdir] = @{
                Count       = 0
                TotalTime   = 0
                Errors      = 0
                SlowCount   = 0
                TimeTaken   = [System.Collections.Generic.List[double]]::new()
            }
        }
        $analysis.VDirStats[$vdir].Count++
        $analysis.VDirStats[$vdir].TotalTime += $entry.TimeTaken
        $analysis.VDirStats[$vdir].TimeTaken.Add($entry.TimeTaken)
        if ($entry.IsError) { $analysis.VDirStats[$vdir].Errors++ }
        if ($entry.IsSlow) { $analysis.VDirStats[$vdir].SlowCount++ }
        
        # Requests per hour
        $hourKey = $entry.DateTime.Substring(0, 13)  # "yyyy-MM-dd HH"
        if (-not $analysis.RequestsPerHour.ContainsKey($hourKey)) { $analysis.RequestsPerHour[$hourKey] = 0 }
        $analysis.RequestsPerHour[$hourKey]++
        
        if ($entry.IsError) {
            if (-not $analysis.ErrorsPerHour.ContainsKey($hourKey)) { $analysis.ErrorsPerHour[$hourKey] = 0 }
            $analysis.ErrorsPerHour[$hourKey]++
        }
        
        # Per-user tracking
        $user = $entry.Username
        if ($user -ne '-' -and $user -ne '') {
            if (-not $userRequests.ContainsKey($user)) {
                $userRequests[$user] = 0
                $userLatency[$user] = [System.Collections.Generic.List[double]]::new()
            }
            $userRequests[$user]++
            $userLatency[$user].Add($entry.TimeTaken)
        }
        
        # Client IP tracking
        $cip = $entry.ClientIP
        if ($cip -ne '-') {
            if (-not $clientIPCount.ContainsKey($cip)) { $clientIPCount[$cip] = 0 }
            $clientIPCount[$cip]++
        }
        
        # Collect auth failures (401)
        if ($entry.StatusCode -eq 401) {
            $analysis.AuthFailures += [PSCustomObject]@{
                DateTime  = $entry.DateTime
                Username  = $entry.Username
                ClientIP  = $entry.ClientIP
                VDir      = $entry.VDir
                UriStem   = $entry.UriStem
            }
        }
        
        # Collect server errors (500+)
        if ($entry.StatusCode -ge 500) {
            $analysis.ServerErrors += [PSCustomObject]@{
                DateTime   = $entry.DateTime
                StatusCode = $entry.StatusCode
                SubStatus  = $entry.SubStatus
                Username   = $entry.Username
                VDir       = $entry.VDir
                UriStem    = $entry.UriStem
                TimeTaken  = $entry.TimeTaken
            }
        }
    }
    
    Write-Progress -Id 2 -ParentId 0 -Activity "Analyzing IIS entries" -Completed
    
    # Calculate overall latency stats
    $sortedTimes = $allTimeTaken | Sort-Object
    $analysis.AvgResponseTime = [math]::Round(($allTimeTaken | Measure-Object -Average).Average, 0)
    $analysis.P95ResponseTime = [math]::Round((Get-PercentileValue -SortedValues $sortedTimes -Percentile 95), 0)
    $analysis.MaxResponseTime = [math]::Round(($allTimeTaken | Measure-Object -Maximum).Maximum, 0)
    
    # Calculate per-vdir P95
    foreach ($vdir in $analysis.VDirStats.Keys) {
        $vdirTimes = $analysis.VDirStats[$vdir].TimeTaken | Sort-Object
        $analysis.VDirStats[$vdir]['AvgTime'] = [math]::Round($analysis.VDirStats[$vdir].TotalTime / $analysis.VDirStats[$vdir].Count, 0)
        $analysis.VDirStats[$vdir]['P95Time'] = [math]::Round((Get-PercentileValue -SortedValues $vdirTimes -Percentile 95), 0)
        $analysis.VDirStats[$vdir]['ErrorRate'] = [math]::Round(($analysis.VDirStats[$vdir].Errors / $analysis.VDirStats[$vdir].Count) * 100, 1)
    }
    
    # Top 20 users by request count
    $analysis.TopUsersByCount = $userRequests.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 20 |
        ForEach-Object { [PSCustomObject]@{ Username = $_.Key; Requests = $_.Value } }
    
    # Top 20 users by average latency (min 10 requests to avoid noise)
    $analysis.TopUsersByLatency = $userLatency.GetEnumerator() |
        Where-Object { $_.Value.Count -ge 10 } |
        ForEach-Object {
            $avg = ($_.Value | Measure-Object -Average).Average
            [PSCustomObject]@{ Username = $_.Key; AvgLatency = [math]::Round($avg, 0); Requests = $_.Value.Count }
        } |
        Sort-Object AvgLatency -Descending |
        Select-Object -First 20
    
    # Top 20 client IPs
    $analysis.TopClientIPs = $clientIPCount.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 20 |
        ForEach-Object { [PSCustomObject]@{ ClientIP = $_.Key; Requests = $_.Value } }
    
    # Slow requests by vdir
    foreach ($vdir in $analysis.VDirStats.Keys) {
        $analysis.SlowByVDir[$vdir] = $analysis.VDirStats[$vdir].SlowCount
    }
    
    # Limit auth failures and server errors to top 50
    $analysis.AuthFailures = $analysis.AuthFailures | Select-Object -First 50
    $analysis.ServerErrors = $analysis.ServerErrors | Select-Object -First 50
    
    return $analysis
}

function Get-HttpProxyAnalysis {
    <#
    .SYNOPSIS
        Analyzes parsed HttpProxy log entries and returns structured analysis results.
    #>
    param(
        [System.Collections.Generic.List[PSObject]]$Entries
    )
    
    $analysis = @{
        TotalRequests         = $Entries.Count
        AvgTotalLatency       = 0
        AvgBackendLatency     = 0
        AvgAuthLatency        = 0
        P95TotalLatency       = 0
        P95BackendLatency     = 0
        BackendServers        = @{}
        ProxyTypeStats        = @{}
        ProxyErrors           = @()
        LatencyByProxyType    = @{}
        RequestsPerHour       = @{}
    }
    
    if ($Entries.Count -eq 0) { return $analysis }
    
    $totalLatencies = [System.Collections.Generic.List[double]]::new()
    $backendLatencies = [System.Collections.Generic.List[double]]::new()
    $authLatencies = [System.Collections.Generic.List[double]]::new()
    
    $entryCount = $Entries.Count
    $processedCount = 0
    
    foreach ($entry in $Entries) {
        $processedCount++
        if ($processedCount % 5000 -eq 0 -or $processedCount -eq $entryCount) {
            $pct = [math]::Round(($processedCount / $entryCount) * 100)
            Write-Progress -Id 2 -ParentId 0 -Activity "Analyzing HttpProxy entries" `
                -Status "$processedCount of $entryCount entries ($pct%)" -PercentComplete $pct
        }
        $totalLatencies.Add($entry.TotalLatency)
        if ($entry.BackendLatency -gt 0) { $backendLatencies.Add($entry.BackendLatency) }
        if ($entry.AuthLatency -gt 0) { $authLatencies.Add($entry.AuthLatency) }
        
        # Backend server distribution
        $backend = $entry.BackendServer
        if ($backend -ne '-' -and $backend -ne '') {
            if (-not $analysis.BackendServers.ContainsKey($backend)) {
                $analysis.BackendServers[$backend] = @{ Count = 0; TotalLatency = 0 }
            }
            $analysis.BackendServers[$backend].Count++
            $analysis.BackendServers[$backend].TotalLatency += $entry.BackendLatency
        }
        
        # Proxy type stats
        $ptype = $entry.ProxyType
        if (-not $analysis.ProxyTypeStats.ContainsKey($ptype)) {
            $analysis.ProxyTypeStats[$ptype] = @{
                Count         = 0
                Errors        = 0
                TotalLatency  = [System.Collections.Generic.List[double]]::new()
                BackendLatency = [System.Collections.Generic.List[double]]::new()
            }
        }
        $analysis.ProxyTypeStats[$ptype].Count++
        $analysis.ProxyTypeStats[$ptype].TotalLatency.Add($entry.TotalLatency)
        if ($entry.BackendLatency -gt 0) { $analysis.ProxyTypeStats[$ptype].BackendLatency.Add($entry.BackendLatency) }
        if ($entry.HttpStatus -ge 400) { $analysis.ProxyTypeStats[$ptype].Errors++ }
        
        # Proxy errors
        if ($entry.ErrorCode -ne '-' -and $entry.ErrorCode -ne '' -and $null -ne $entry.ErrorCode) {
            $analysis.ProxyErrors += [PSCustomObject]@{
                DateTime          = $entry.DateTime
                ProxyType         = $entry.ProxyType
                ErrorCode         = $entry.ErrorCode
                GenericErrors     = $entry.GenericErrors
                HttpStatus        = $entry.HttpStatus
                AuthenticatedUser = $entry.AuthenticatedUser
                BackendServer     = $entry.BackendServer
                TotalLatency      = $entry.TotalLatency
            }
        }
        
        # Requests per hour
        $hourKey = $entry.DateTime.ToString('yyyy-MM-dd HH')
        if (-not $analysis.RequestsPerHour.ContainsKey($hourKey)) { $analysis.RequestsPerHour[$hourKey] = 0 }
        $analysis.RequestsPerHour[$hourKey]++
    }
    
    Write-Progress -Id 2 -ParentId 0 -Activity "Analyzing HttpProxy entries" -Completed
    
    # Overall latency stats
    if ($totalLatencies.Count -gt 0) {
        $analysis.AvgTotalLatency = [math]::Round(($totalLatencies | Measure-Object -Average).Average, 0)
        $sorted = $totalLatencies | Sort-Object
        $analysis.P95TotalLatency = [math]::Round((Get-PercentileValue -SortedValues $sorted -Percentile 95), 0)
    }
    if ($backendLatencies.Count -gt 0) {
        $analysis.AvgBackendLatency = [math]::Round(($backendLatencies | Measure-Object -Average).Average, 0)
        $sorted = $backendLatencies | Sort-Object
        $analysis.P95BackendLatency = [math]::Round((Get-PercentileValue -SortedValues $sorted -Percentile 95), 0)
    }
    if ($authLatencies.Count -gt 0) {
        $analysis.AvgAuthLatency = [math]::Round(($authLatencies | Measure-Object -Average).Average, 0)
    }
    
    # Calculate per proxy type averages
    foreach ($ptype in $analysis.ProxyTypeStats.Keys) {
        $stats = $analysis.ProxyTypeStats[$ptype]
        $sortedTotal = $stats.TotalLatency | Sort-Object
        
        $stats['AvgTotalLatency'] = if ($stats.TotalLatency.Count -gt 0) {
            [math]::Round(($stats.TotalLatency | Measure-Object -Average).Average, 0)
        } else { 0 }
        $stats['P95TotalLatency'] = if ($sortedTotal.Count -gt 0) {
            [math]::Round((Get-PercentileValue -SortedValues $sortedTotal -Percentile 95), 0)
        } else { 0 }
        $stats['AvgBackendLatency'] = if ($stats.BackendLatency.Count -gt 0) {
            [math]::Round(($stats.BackendLatency | Measure-Object -Average).Average, 0)
        } else { 0 }
        $stats['ErrorRate'] = if ($stats.Count -gt 0) {
            [math]::Round(($stats.Errors / $stats.Count) * 100, 1)
        } else { 0 }
    }
    
    # Calculate backend server averages
    foreach ($backend in $analysis.BackendServers.Keys) {
        $bs = $analysis.BackendServers[$backend]
        $bs['AvgLatency'] = if ($bs.Count -gt 0) { [math]::Round($bs.TotalLatency / $bs.Count, 0) } else { 0 }
    }
    
    # Limit proxy errors to top 100
    $analysis.ProxyErrors = $analysis.ProxyErrors | Select-Object -First 100
    
    return $analysis
}

#endregion

#region HTML Report Generation

function New-IISHTMLReport {
    <#
    .SYNOPSIS
        Generates an HTML report from the analysis results.
    #>
    param(
        [hashtable]$ReportData
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $hoursAnalyzed = $ReportData.Hours
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange IIS Log Analysis Report</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f0f2f5; color: #333; }
        .header { background: linear-gradient(135deg, #1a5276, #2980b9); color: white; padding: 30px; }
        .header h1 { font-size: 24px; margin-bottom: 5px; }
        .header p { opacity: 0.9; font-size: 14px; }
        .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .summary-card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; }
        .summary-card .value { font-size: 28px; font-weight: bold; margin: 5px 0; }
        .summary-card .label { font-size: 12px; text-transform: uppercase; color: #666; letter-spacing: 1px; }
        .server-section { background: white; border-radius: 8px; padding: 25px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .server-section h2 { color: #1a5276; border-left: 4px solid #2980b9; padding-left: 15px; margin-bottom: 20px; }
        .server-section h3 { color: #2c3e50; margin: 20px 0 10px 0; font-size: 16px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 13px; }
        th { background: #2c3e50; color: white; padding: 10px 12px; text-align: left; font-weight: 600; }
        td { padding: 8px 12px; border-bottom: 1px solid #eee; }
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e8f4fd; }
        .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 15px; margin: 15px 0; }
        .metric-card { background: #f8f9fa; border-radius: 6px; padding: 15px; border-left: 4px solid #2980b9; }
        .metric-card .metric-label { font-size: 11px; text-transform: uppercase; color: #666; letter-spacing: 0.5px; }
        .metric-card .metric-value { font-size: 22px; font-weight: bold; margin: 5px 0; }
        .status-healthy { color: #27ae60; }
        .status-warning { color: #f39c12; }
        .status-critical { color: #e74c3c; }
        .status-info { color: #2980b9; }
        .status-notavailable { color: #95a5a6; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: 600; }
        .badge-error { background: #fde8e8; color: #e74c3c; }
        .badge-warning { background: #fef5e7; color: #f39c12; }
        .badge-success { background: #e8f8f5; color: #27ae60; }
        .badge-info { background: #ebf5fb; color: #2980b9; }
        .error-list { max-height: 400px; overflow-y: auto; }
        .footer { text-align: center; padding: 20px; color: #666; font-size: 12px; }
        .tab-container { margin: 15px 0; }
        .vdir-bar { display: flex; align-items: center; margin: 4px 0; }
        .vdir-bar .bar-label { width: 200px; font-size: 13px; }
        .vdir-bar .bar-track { flex: 1; height: 24px; background: #ecf0f1; border-radius: 4px; overflow: hidden; }
        .vdir-bar .bar-fill { height: 100%; border-radius: 4px; display: flex; align-items: center; padding-left: 8px; font-size: 11px; color: white; font-weight: 600; }
        .bar-healthy { background: #27ae60; }
        .bar-warning { background: #f39c12; }
        .bar-error { background: #e74c3c; }
        .bar-info { background: #3498db; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
        @media (max-width: 768px) { .two-col { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="header">
        <h1>Exchange IIS & HttpProxy Log Analysis</h1>
        <p>Generated: $timestamp | Analysis Period: Last $hoursAnalyzed hours | Servers: $($ReportData.Servers -join ', ')</p>
    </div>
    <div class="container">
"@

    # Per-server sections
    foreach ($serverName in $ReportData.Servers) {
        $serverData = $ReportData.ServerResults[$serverName]
        
        $html += @"
        <div class="server-section">
            <h2>Server: $serverName</h2>
"@
        
        # IIS Analysis
        if ($serverData.IIS) {
            $iis = $serverData.IIS
            
            # Summary cards
            $errorRateColor = if (($iis.ErrorRequests / [math]::Max($iis.TotalRequests, 1) * 100) -gt 5) { 'critical' } elseif (($iis.ErrorRequests / [math]::Max($iis.TotalRequests, 1) * 100) -gt 1) { 'warning' } else { 'healthy' }
            $p95Color = if ($iis.P95ResponseTime -gt $ReportData.SlowThreshold) { 'critical' } elseif ($iis.P95ResponseTime -gt ($ReportData.SlowThreshold / 2)) { 'warning' } else { 'healthy' }
            $slowColor = if ($iis.SlowRequests -gt 100) { 'critical' } elseif ($iis.SlowRequests -gt 10) { 'warning' } else { 'healthy' }
            
            $html += @"
            <h3>IIS Log Summary</h3>
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Total Requests</div>
                    <div class="metric-value status-info">$($iis.TotalRequests.ToString('N0'))</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Avg Response Time</div>
                    <div class="metric-value status-$p95Color">$($iis.AvgResponseTime) ms</div>
                    <div class="metric-label">P95: $($iis.P95ResponseTime) ms | Max: $($iis.MaxResponseTime) ms</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Error Requests (4xx/5xx)</div>
                    <div class="metric-value status-$errorRateColor">$($iis.ErrorRequests.ToString('N0'))</div>
                    <div class="metric-label">$([math]::Round($iis.ErrorRequests / [math]::Max($iis.TotalRequests, 1) * 100, 1))% error rate</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Slow Requests (>$($ReportData.SlowThreshold)ms)</div>
                    <div class="metric-value status-$slowColor">$($iis.SlowRequests.ToString('N0'))</div>
                </div>
            </div>
"@
            
            # Virtual Directory Breakdown
            if ($iis.VDirStats.Count -gt 0) {
                $html += @"
            <h3>Virtual Directory Breakdown</h3>
            <table>
                <tr><th>Virtual Directory</th><th>Requests</th><th>Avg (ms)</th><th>P95 (ms)</th><th>Errors</th><th>Error Rate</th><th>Slow</th></tr>
"@
                foreach ($vdir in ($iis.VDirStats.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending)) {
                    $vdirName = Get-FriendlyVDirName -VDir $vdir.Key
                    $errRateClass = if ($vdir.Value.ErrorRate -gt 5) { 'status-critical' } elseif ($vdir.Value.ErrorRate -gt 1) { 'status-warning' } else { 'status-healthy' }
                    $avgClass = if ($vdir.Value.AvgTime -gt $ReportData.SlowThreshold) { 'status-critical' } elseif ($vdir.Value.AvgTime -gt ($ReportData.SlowThreshold / 2)) { 'status-warning' } else { 'status-healthy' }
                    
                    $html += @"
                <tr>
                    <td><strong>$vdirName</strong><br><small>/$($vdir.Key)/</small></td>
                    <td>$($vdir.Value.Count.ToString('N0'))</td>
                    <td class="$avgClass">$($vdir.Value.AvgTime)</td>
                    <td class="$avgClass">$($vdir.Value.P95Time)</td>
                    <td>$($vdir.Value.Errors.ToString('N0'))</td>
                    <td class="$errRateClass">$($vdir.Value.ErrorRate)%</td>
                    <td>$($vdir.Value.SlowCount)</td>
                </tr>
"@
                }
                $html += "            </table>"
            }
            
            # HTTP Status Code Distribution
            if ($iis.StatusCodes.Count -gt 0) {
                $html += @"
            <h3>HTTP Status Code Distribution</h3>
            <table>
                <tr><th>Status Code</th><th>Count</th><th>Percentage</th><th>Description</th></tr>
"@
                $statusDescriptions = @{
                    '200' = 'OK'; '301' = 'Moved Permanently'; '302' = 'Found (Redirect)'
                    '304' = 'Not Modified'; '401' = 'Unauthorized'; '403' = 'Forbidden'
                    '404' = 'Not Found'; '408' = 'Request Timeout'; '500' = 'Internal Server Error'
                    '502' = 'Bad Gateway'; '503' = 'Service Unavailable'; '504' = 'Gateway Timeout'
                }
                
                foreach ($sc in ($iis.StatusCodes.GetEnumerator() | Sort-Object Value -Descending)) {
                    $pct = [math]::Round(($sc.Value / $iis.TotalRequests) * 100, 2)
                    $desc = if ($statusDescriptions.ContainsKey($sc.Key)) { $statusDescriptions[$sc.Key] } else { '-' }
                    $scClass = if ([int]$sc.Key -ge 500) { 'status-critical' } elseif ([int]$sc.Key -ge 400) { 'status-warning' } else { 'status-healthy' }
                    
                    $html += @"
                <tr>
                    <td class="$scClass"><strong>$($sc.Key)</strong></td>
                    <td>$($sc.Value.ToString('N0'))</td>
                    <td>$pct%</td>
                    <td>$desc</td>
                </tr>
"@
                }
                $html += "            </table>"
            }
            
            # Top Talkers (two-column)
            $html += @"
            <div class="two-col">
                <div>
                    <h3>Top 20 Users by Request Count</h3>
                    <table>
                        <tr><th>#</th><th>Username</th><th>Requests</th></tr>
"@
            $rank = 1
            foreach ($user in $iis.TopUsersByCount) {
                $html += "                        <tr><td>$rank</td><td>$($user.Username)</td><td>$($user.Requests.ToString('N0'))</td></tr>`n"
                $rank++
            }
            
            $html += @"
                    </table>
                </div>
                <div>
                    <h3>Top 20 Users by Avg Latency</h3>
                    <table>
                        <tr><th>#</th><th>Username</th><th>Avg (ms)</th><th>Requests</th></tr>
"@
            $rank = 1
            foreach ($user in $iis.TopUsersByLatency) {
                $latClass = if ($user.AvgLatency -gt $ReportData.SlowThreshold) { 'status-critical' } elseif ($user.AvgLatency -gt ($ReportData.SlowThreshold / 2)) { 'status-warning' } else { '' }
                $html += "                        <tr><td>$rank</td><td>$($user.Username)</td><td class='$latClass'>$($user.AvgLatency)</td><td>$($user.Requests)</td></tr>`n"
                $rank++
            }
            
            $html += @"
                    </table>
                </div>
            </div>
"@

            # Top Client IPs
            if ($iis.TopClientIPs.Count -gt 0) {
                $html += @"
            <h3>Top 20 Client IPs</h3>
            <table>
                <tr><th>#</th><th>Client IP</th><th>Requests</th></tr>
"@
                $rank = 1
                foreach ($cip in $iis.TopClientIPs) {
                    $html += "                <tr><td>$rank</td><td>$($cip.ClientIP)</td><td>$($cip.Requests.ToString('N0'))</td></tr>`n"
                    $rank++
                }
                $html += "            </table>"
            }
            
            # Authentication Failures
            if ($iis.AuthFailures.Count -gt 0) {
                $html += @"
            <h3>Authentication Failures (401) <span class="badge badge-error">$($iis.AuthFailures.Count) shown</span></h3>
            <div class="error-list">
            <table>
                <tr><th>Time</th><th>Username</th><th>Client IP</th><th>Virtual Directory</th></tr>
"@
                foreach ($af in $iis.AuthFailures) {
                    $html += "                <tr><td>$($af.DateTime)</td><td>$($af.Username)</td><td>$($af.ClientIP)</td><td>$($af.VDir)</td></tr>`n"
                }
                $html += "            </table></div>"
            }
            
            # Server Errors
            if ($iis.ServerErrors.Count -gt 0) {
                $html += @"
            <h3>Server Errors (5xx) <span class="badge badge-error">$($iis.ServerErrors.Count) shown</span></h3>
            <div class="error-list">
            <table>
                <tr><th>Time</th><th>Status</th><th>Username</th><th>VDir</th><th>URI</th><th>Time Taken</th></tr>
"@
                foreach ($se in $iis.ServerErrors) {
                    $html += "                <tr><td>$($se.DateTime)</td><td class='status-critical'>$($se.StatusCode).$($se.SubStatus)</td><td>$($se.Username)</td><td>$($se.VDir)</td><td style='max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'>$($se.UriStem)</td><td>$($se.TimeTaken)ms</td></tr>`n"
                }
                $html += "            </table></div>"
            }
        }
        
        # HttpProxy Analysis
        if ($serverData.HttpProxy) {
            $proxy = $serverData.HttpProxy
            
            $html += @"
            <h3>HttpProxy Log Analysis</h3>
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Total Proxy Requests</div>
                    <div class="metric-value status-info">$($proxy.TotalRequests.ToString('N0'))</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Avg Total Latency</div>
                    <div class="metric-value">$($proxy.AvgTotalLatency) ms</div>
                    <div class="metric-label">P95: $($proxy.P95TotalLatency) ms</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Avg Backend Latency</div>
                    <div class="metric-value">$($proxy.AvgBackendLatency) ms</div>
                    <div class="metric-label">P95: $($proxy.P95BackendLatency) ms</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Avg Auth Latency</div>
                    <div class="metric-value">$($proxy.AvgAuthLatency) ms</div>
                </div>
            </div>
"@
            
            # Proxy Type Stats
            if ($proxy.ProxyTypeStats.Count -gt 0) {
                $html += @"
            <h3>Latency by Proxy Type</h3>
            <table>
                <tr><th>Proxy Type</th><th>Requests</th><th>Avg Total (ms)</th><th>P95 Total (ms)</th><th>Avg Backend (ms)</th><th>Errors</th><th>Error Rate</th></tr>
"@
                foreach ($pt in ($proxy.ProxyTypeStats.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending)) {
                    $errClass = if ($pt.Value.ErrorRate -gt 5) { 'status-critical' } elseif ($pt.Value.ErrorRate -gt 1) { 'status-warning' } else { 'status-healthy' }
                    
                    $html += @"
                <tr>
                    <td><strong>$($pt.Key)</strong></td>
                    <td>$($pt.Value.Count.ToString('N0'))</td>
                    <td>$($pt.Value.AvgTotalLatency)</td>
                    <td>$($pt.Value.P95TotalLatency)</td>
                    <td>$($pt.Value.AvgBackendLatency)</td>
                    <td>$($pt.Value.Errors)</td>
                    <td class="$errClass">$($pt.Value.ErrorRate)%</td>
                </tr>
"@
                }
                $html += "            </table>"
            }
            
            # Backend Server Distribution
            if ($proxy.BackendServers.Count -gt 0) {
                $html += @"
            <h3>Backend Server Distribution</h3>
            <table>
                <tr><th>Backend Server</th><th>Requests</th><th>Avg Backend Latency (ms)</th></tr>
"@
                foreach ($bs in ($proxy.BackendServers.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | Select-Object -First 20)) {
                    $latClass = if ($bs.Value.AvgLatency -gt 5000) { 'status-critical' } elseif ($bs.Value.AvgLatency -gt 2000) { 'status-warning' } else { '' }
                    $html += "                <tr><td>$($bs.Key)</td><td>$($bs.Value.Count.ToString('N0'))</td><td class='$latClass'>$($bs.Value.AvgLatency)</td></tr>`n"
                }
                $html += "            </table>"
            }
            
            # Proxy Errors
            if ($proxy.ProxyErrors.Count -gt 0) {
                $html += @"
            <h3>Proxy Errors <span class="badge badge-error">$($proxy.ProxyErrors.Count) shown</span></h3>
            <div class="error-list">
            <table>
                <tr><th>Time</th><th>Type</th><th>Error Code</th><th>HTTP Status</th><th>User</th><th>Backend</th><th>Latency</th></tr>
"@
                foreach ($pe in $proxy.ProxyErrors) {
                    $html += "                <tr><td>$($pe.DateTime)</td><td>$($pe.ProxyType)</td><td class='status-critical'>$($pe.ErrorCode)</td><td>$($pe.HttpStatus)</td><td>$($pe.AuthenticatedUser)</td><td>$($pe.BackendServer)</td><td>$($pe.TotalLatency)ms</td></tr>`n"
                }
                $html += "            </table></div>"
            }
        }
        
        $html += "        </div>"  # Close server-section
    }
    
    # Footer
    $html += @"
    </div>
    <div class="footer">
        <p>Exchange IIS & HttpProxy Log Analysis | Generated by Analyze-ExchangeIISLogs.ps1 | $timestamp</p>
    </div>
</body>
</html>
"@
    
    return $html
}

#endregion

#region Main Execution

Write-AnalysisLog "========== Exchange IIS & HttpProxy Log Analyzer ==========" -Level Info

# RunLocally mode: force to local server and use local paths
if ($RunLocally) {
    $Servers = @($env:COMPUTERNAME)
    Write-AnalysisLog "Mode: LOCAL (reading logs from local disk - fastest)" -Level Success
} else {
    Write-AnalysisLog "Mode: REMOTE (reading logs via UNC paths)" -Level Info
}

Write-AnalysisLog "Servers: $($Servers -join ', ')" -Level Info
Write-AnalysisLog "Time range: Last $Hours hours" -Level Info
Write-AnalysisLog "Slow request threshold: ${SlowRequestThreshold}ms" -Level Info
Write-AnalysisLog "IIS sites: $LogSite" -Level Info
Write-AnalysisLog "HttpProxy analysis: $(if ($IncludeHttpProxy) { 'Enabled' } else { 'Disabled' })" -Level Info

$startTime = (Get-Date).AddHours(-$Hours)
$reportData = @{
    Servers        = $Servers
    Hours          = $Hours
    SlowThreshold  = $SlowRequestThreshold
    ServerResults  = @{}
}

# Determine which IIS sites to scan
$iisSites = @()
if (-not $SkipIISLogs) {
    switch ($LogSite) {
        'Frontend' { $iisSites = @('W3SVC1') }
        'Backend'  { $iisSites = @('W3SVC2') }
        'Both'     { $iisSites = @('W3SVC1', 'W3SVC2') }
    }
}

# HttpProxy types to scan
$httpProxyTypes = @('Mapi', 'Ews', 'OwaProxy', 'RpcHttp', 'Autodiscover', 'ActiveSync')

# Progress tracking: calculate total steps for percentage
# Phase 1 (0-15%): Start jobs  |  Phase 2 (15-60%): Wait for jobs  |  Phase 3 (60-85%): Analyze  |  Phase 4 (85-100%): Report
$totalServers = $Servers.Count
$serverIndex = 0

# Step 1: Start background jobs for all servers (log parsing is I/O bound)
$iisJobs = @{}
$proxyJobs = @{}

foreach ($serverName in $Servers) {
    $serverIndex++
    $overallPct = [math]::Round(($serverIndex / $totalServers) * 15)
    Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
        -Status "Phase 1/4: Starting jobs - Server $serverIndex of $totalServers ($serverName)" -PercentComplete $overallPct

    Write-AnalysisLog "`n--- Starting log collection for $serverName ---" -Level Info
    $iisJobs[$serverName] = @{}
    $proxyJobs[$serverName] = @{}
    
    # Start IIS log parsing jobs
    foreach ($site in $iisSites) {
        $iisJobs[$serverName][$site] = Start-Job -ScriptBlock {
            param($Server, $Site, $Start, $Threshold, $Local)
            
            $logPath = if ($Local) { "C:\inetpub\logs\LogFiles\$Site" } else { "\\$Server\C`$\inetpub\logs\LogFiles\$Site" }
            $result = @{
                Entries        = @()
                TotalLines     = 0
                ParseErrors    = 0
                FilesProcessed = 0
                Status         = 'Unknown'
                Site           = $Site
            }
            
            if (-not (Test-Path -Path $logPath -PathType Container)) {
                $result.Status = 'PathNotFound'
                return $result
            }
            
            $logFiles = Get-ChildItem -Path $logPath -Filter '*.log' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $Start.AddHours(-1) } |
                Sort-Object LastWriteTime
            
            if (-not $logFiles) {
                $result.Status = 'NoFilesFound'
                return $result
            }
            
            $startDateStr = $Start.ToString('yyyy-MM-dd')
            $startTimeStr = $Start.ToString('HH:mm:ss')
            $entries = [System.Collections.Generic.List[PSObject]]::new()
            
            foreach ($logFile in $logFiles) {
                $result.FilesProcessed++
                $fields = $null
                
                try {
                    $reader = [System.IO.StreamReader]::new($logFile.FullName)
                    
                    while (-not $reader.EndOfStream) {
                        $line = $reader.ReadLine()
                        $result.TotalLines++
                        
                        if ($line.StartsWith('#Fields:')) {
                            if ($line -match '^#Fields:\s+(.+)$') {
                                $fields = ($Matches[1] -split '\s+')
                            }
                            continue
                        }
                        
                        if ($line.StartsWith('#')) { continue }
                        if ($null -eq $fields) { continue }
                        
                        $lineParts = $line -split '\s+', 3
                        if ($lineParts.Count -lt 2) { continue }
                        if ($lineParts[0] -lt $startDateStr) { continue }
                        if ($lineParts[0] -eq $startDateStr -and $lineParts[1] -lt $startTimeStr) { continue }
                        
                        $values = $line -split '\s+'
                        if ($values.Count -ne $fields.Count) {
                            $result.ParseErrors++
                            continue
                        }
                        
                        $entry = @{}
                        for ($i = 0; $i -lt $fields.Count; $i++) {
                            $entry[$fields[$i]] = $values[$i]
                        }
                        
                        $timeTaken = 0
                        if ($entry.ContainsKey('time-taken')) {
                            [int]::TryParse($entry['time-taken'], [ref]$timeTaken) | Out-Null
                        }
                        
                        $statusCode = 0
                        if ($entry.ContainsKey('sc-status')) {
                            [int]::TryParse($entry['sc-status'], [ref]$statusCode) | Out-Null
                        }
                        
                        $uriStem = if ($entry.ContainsKey('cs-uri-stem')) { $entry['cs-uri-stem'] } else { '/' }
                        
                        # Determine VDir
                        $vdir = 'Other'
                        $knownVDirs = @('/mapi/', '/rpc/', '/owa/', '/ews/', '/Microsoft-Server-ActiveSync/', '/autodiscover/', '/oab/', '/powershell/', '/ecp/')
                        foreach ($kv in $knownVDirs) {
                            if ($uriStem -like "$kv*" -or $uriStem -eq $kv.TrimEnd('/')) {
                                $vdir = $kv.Trim('/')
                                break
                            }
                        }
                        
                        $parsedEntry = [PSCustomObject]@{
                            DateTime   = "$($entry['date']) $($entry['time'])"
                            Method     = if ($entry.ContainsKey('cs-method')) { $entry['cs-method'] } else { '-' }
                            UriStem    = $uriStem
                            VDir       = $vdir
                            StatusCode = $statusCode
                            SubStatus  = if ($entry.ContainsKey('sc-substatus')) { $entry['sc-substatus'] } else { '0' }
                            TimeTaken  = $timeTaken
                            ClientIP   = if ($entry.ContainsKey('c-ip')) { $entry['c-ip'] } else { '-' }
                            Username   = if ($entry.ContainsKey('cs-username')) { $entry['cs-username'] } else { '-' }
                            UserAgent  = if ($entry.ContainsKey('cs(User-Agent)')) { $entry['cs(User-Agent)'] } else { '-' }
                            ServerIP   = if ($entry.ContainsKey('s-ip')) { $entry['s-ip'] } else { '-' }
                            ServerPort = if ($entry.ContainsKey('s-port')) { $entry['s-port'] } else { '-' }
                            BytesSent  = if ($entry.ContainsKey('sc-bytes')) { $entry['sc-bytes'] } else { '0' }
                            BytesRecv  = if ($entry.ContainsKey('cs-bytes')) { $entry['cs-bytes'] } else { '0' }
                            IsSlow     = ($timeTaken -ge $Threshold)
                            IsError    = ($statusCode -ge 400)
                        }
                        
                        $entries.Add($parsedEntry)
                    }
                    
                    $reader.Close()
                    $reader.Dispose()
                } catch {
                    if ($reader) { $reader.Dispose() }
                    $result.ParseErrors++
                }
            }
            
            $result.Entries = $entries.ToArray()
            $result.Status = 'Collected'
            return $result
            
        } -ArgumentList $serverName, $site, $startTime, $SlowRequestThreshold, $RunLocally.IsPresent
        
        Write-AnalysisLog "  Started IIS $site log parsing job for $serverName$(if ($RunLocally) { ' (local)' })" -Level Info
    }
    
    # Start HttpProxy log parsing jobs
    if ($IncludeHttpProxy) {
        foreach ($proxyType in $httpProxyTypes) {
            $proxyJobs[$serverName][$proxyType] = Start-Job -ScriptBlock {
                param($Server, $PType, $Start, $Local)
                
                $basePath = if ($Local) { "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\$PType" } else { "\\$Server\C`$\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\$PType" }
                $result = @{
                    Entries        = @()
                    TotalLines     = 0
                    ParseErrors    = 0
                    FilesProcessed = 0
                    Status         = 'Unknown'
                    ProxyType      = $PType
                }
                
                if (-not (Test-Path -Path $basePath -PathType Container)) {
                    $result.Status = 'PathNotFound'
                    return $result
                }
                
                $logFiles = Get-ChildItem -Path $basePath -Filter '*.log' -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -ge $Start.AddHours(-1) } |
                    Sort-Object LastWriteTime
                
                if (-not $logFiles) {
                    $result.Status = 'NoFilesFound'
                    return $result
                }
                
                $entries = [System.Collections.Generic.List[PSObject]]::new()
                
                foreach ($logFile in $logFiles) {
                    $result.FilesProcessed++
                    $headers = $null
                    
                    try {
                        $reader = [System.IO.StreamReader]::new($logFile.FullName)
                        
                        while (-not $reader.EndOfStream) {
                            $line = $reader.ReadLine()
                            $result.TotalLines++
                            
                            if ($line.StartsWith('#') -and $line -match 'DateTime') {
                                $headers = ($line.TrimStart('#') -split ',')
                                continue
                            }
                            
                            if ($line.StartsWith('#')) { continue }
                            if ($null -eq $headers) { continue }
                            
                            $values = $line -split ','
                            if ($values.Count -lt $headers.Count) {
                                $result.ParseErrors++
                                continue
                            }
                            
                            $entry = @{}
                            for ($i = 0; $i -lt $headers.Count; $i++) {
                                $entry[$headers[$i].Trim()] = $values[$i].Trim()
                            }
                            
                            $dateTimeStr = if ($entry.ContainsKey('DateTime')) { $entry['DateTime'] } else { $null }
                            if ($null -eq $dateTimeStr) { continue }
                            
                            $entryTime = [datetime]::MinValue
                            if (-not [datetime]::TryParse($dateTimeStr, [ref]$entryTime)) { continue }
                            if ($entryTime -lt $Start) { continue }
                            
                            $totalLatency = 0
                            $backendLatency = 0
                            $authLatency = 0
                            $statusCode = 0
                            
                            if ($entry.ContainsKey('TotalRequestTime')) { [int]::TryParse($entry['TotalRequestTime'], [ref]$totalLatency) | Out-Null }
                            if ($entry.ContainsKey('BackendProcessingLatency')) { [int]::TryParse($entry['BackendProcessingLatency'], [ref]$backendLatency) | Out-Null }
                            if ($entry.ContainsKey('AuthModuleLatency')) { [int]::TryParse($entry['AuthModuleLatency'], [ref]$authLatency) | Out-Null }
                            if ($entry.ContainsKey('HttpStatus')) { [int]::TryParse($entry['HttpStatus'], [ref]$statusCode) | Out-Null }
                            
                            $parsedEntry = [PSCustomObject]@{
                                DateTime              = $entryTime
                                ProxyType             = $PType
                                AuthenticatedUser     = if ($entry.ContainsKey('AuthenticatedUser')) { $entry['AuthenticatedUser'] } else { '-' }
                                UrlStem               = if ($entry.ContainsKey('UrlStem')) { $entry['UrlStem'] } else { '-' }
                                HttpStatus            = $statusCode
                                BackendServer         = if ($entry.ContainsKey('TargetServer')) { $entry['TargetServer'] } else { '-' }
                                TotalLatency          = $totalLatency
                                BackendLatency        = $backendLatency
                                AuthLatency           = $authLatency
                                ClientIP              = if ($entry.ContainsKey('ClientIpAddress')) { $entry['ClientIpAddress'] } else { '-' }
                                ErrorCode             = if ($entry.ContainsKey('ErrorCode')) { $entry['ErrorCode'] } else { '-' }
                                GenericErrors         = if ($entry.ContainsKey('GenericErrors')) { $entry['GenericErrors'] } else { '-' }
                                ServerLocatorLatency  = if ($entry.ContainsKey('ServerLocatorLatency')) { $entry['ServerLocatorLatency'] } else { '0' }
                                TargetServerVersion   = if ($entry.ContainsKey('TargetServerVersion')) { $entry['TargetServerVersion'] } else { '-' }
                                RequestId             = if ($entry.ContainsKey('RequestId')) { $entry['RequestId'] } else { '-' }
                            }
                            
                            $entries.Add($parsedEntry)
                        }
                        
                        $reader.Close()
                        $reader.Dispose()
                    } catch {
                        if ($reader) { $reader.Dispose() }
                        $result.ParseErrors++
                    }
                }
                
                $result.Entries = $entries.ToArray()
                $result.Status = 'Collected'
                return $result
                
            } -ArgumentList $serverName, $proxyType, $startTime, $RunLocally.IsPresent
            
            Write-AnalysisLog "  Started HttpProxy $proxyType log parsing job for $serverName$(if ($RunLocally) { ' (local)' })" -Level Info
        }
    }
}

# Step 2: Wait for all jobs and collect results
Write-AnalysisLog "`n========== Waiting for log parsing jobs to complete ==========" -Level Info

$serverIndex = 0
foreach ($serverName in $Servers) {
    $serverIndex++
    # Phase 2-3 spans 15-85% overall; split evenly per server
    $serverBasePct = 15 + [math]::Round((($serverIndex - 1) / $totalServers) * 70)
    $serverEndPct  = 15 + [math]::Round(($serverIndex / $totalServers) * 70)
    
    Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
        -Status "Phase 2/4: Collecting results - Server $serverIndex of $totalServers ($serverName)" -PercentComplete $serverBasePct

    Write-AnalysisLog "`n--- Processing results for $serverName ---" -Level Info
    
    $reportData.ServerResults[$serverName] = @{
        IIS       = $null
        HttpProxy = $null
    }
    
    # Collect IIS results
    if (-not $SkipIISLogs) {
        $allIISEntries = [System.Collections.Generic.List[PSObject]]::new()
        
        $siteIndex = 0
        foreach ($site in $iisSites) {
            $siteIndex++
            if ($iisJobs[$serverName].ContainsKey($site)) {
                Write-Progress -Id 1 -ParentId 0 -Activity "Waiting for IIS jobs on $serverName" `
                    -Status "Site $siteIndex of $($iisSites.Count): $site" -PercentComplete ([math]::Round(($siteIndex / $iisSites.Count) * 100))
                Write-AnalysisLog "  Waiting for IIS $site job on $serverName..." -Level Info
                $jobResult = Receive-Job -Job $iisJobs[$serverName][$site] -Wait
                Remove-Job -Job $iisJobs[$serverName][$site] -Force
                
                if ($jobResult.Status -eq 'Collected') {
                    Write-AnalysisLog "  IIS ${site}: $($jobResult.Entries.Count) entries from $($jobResult.FilesProcessed) files ($($jobResult.TotalLines) lines parsed)" -Level Success
                    foreach ($entry in $jobResult.Entries) {
                        $allIISEntries.Add($entry)
                    }
                } elseif ($jobResult.Status -eq 'PathNotFound') {
                    Write-AnalysisLog "  IIS ${site}: Log path not accessible on $serverName" -Level Warning
                } elseif ($jobResult.Status -eq 'NoFilesFound') {
                    Write-AnalysisLog "  IIS ${site}: No log files found in time range on $serverName" -Level Warning
                } else {
                    Write-AnalysisLog "  IIS ${site}: Unknown status - $($jobResult.Status)" -Level Warning
                }
            }
        }
        
        Write-Progress -Id 1 -ParentId 0 -Activity "Waiting for IIS jobs on $serverName" -Completed
        
        if ($allIISEntries.Count -gt 0) {
            $analysisPct = $serverBasePct + [math]::Round(($serverEndPct - $serverBasePct) * 0.5)
            Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
                -Status "Phase 3/4: Analyzing IIS entries for $serverName ($($allIISEntries.Count) entries)" -PercentComplete $analysisPct
            Write-AnalysisLog "  Analyzing $($allIISEntries.Count) IIS log entries for $serverName..." -Level Info
            $reportData.ServerResults[$serverName].IIS = Get-IISAnalysis -Entries $allIISEntries -SlowThreshold $SlowRequestThreshold
            Write-AnalysisLog "  IIS analysis complete: $($reportData.ServerResults[$serverName].IIS.TotalRequests) requests, $($reportData.ServerResults[$serverName].IIS.ErrorRequests) errors, $($reportData.ServerResults[$serverName].IIS.SlowRequests) slow" -Level Success
        } else {
            Write-AnalysisLog "  No IIS log entries collected for $serverName" -Level Warning
        }
    }
    
    # Collect HttpProxy results
    if ($IncludeHttpProxy) {
        $allProxyEntries = [System.Collections.Generic.List[PSObject]]::new()
        
        $proxyIndex = 0
        foreach ($proxyType in $httpProxyTypes) {
            $proxyIndex++
            if ($proxyJobs[$serverName].ContainsKey($proxyType)) {
                Write-Progress -Id 1 -ParentId 0 -Activity "Waiting for HttpProxy jobs on $serverName" `
                    -Status "Type $proxyIndex of $($httpProxyTypes.Count): $proxyType" -PercentComplete ([math]::Round(($proxyIndex / $httpProxyTypes.Count) * 100))
                Write-AnalysisLog "  Waiting for HttpProxy $proxyType job on $serverName..." -Level Info
                $jobResult = Receive-Job -Job $proxyJobs[$serverName][$proxyType] -Wait
                Remove-Job -Job $proxyJobs[$serverName][$proxyType] -Force
                
                if ($jobResult.Status -eq 'Collected') {
                    Write-AnalysisLog "  HttpProxy ${proxyType}: $($jobResult.Entries.Count) entries from $($jobResult.FilesProcessed) files" -Level Success
                    foreach ($entry in $jobResult.Entries) {
                        $allProxyEntries.Add($entry)
                    }
                } elseif ($jobResult.Status -eq 'PathNotFound') {
                    Write-AnalysisLog "  HttpProxy ${proxyType}: Path not accessible on $serverName" -Level Warning
                } elseif ($jobResult.Status -eq 'NoFilesFound') {
                    Write-AnalysisLog "  HttpProxy ${proxyType}: No log files found in time range" -Level Warning
                }
            }
        }
        
        Write-Progress -Id 1 -ParentId 0 -Activity "Waiting for HttpProxy jobs on $serverName" -Completed
        
        if ($allProxyEntries.Count -gt 0) {
            $analysisPct = $serverBasePct + [math]::Round(($serverEndPct - $serverBasePct) * 0.8)
            Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
                -Status "Phase 3/4: Analyzing HttpProxy entries for $serverName ($($allProxyEntries.Count) entries)" -PercentComplete $analysisPct
            Write-AnalysisLog "  Analyzing $($allProxyEntries.Count) HttpProxy entries for $serverName..." -Level Info
            $reportData.ServerResults[$serverName].HttpProxy = Get-HttpProxyAnalysis -Entries $allProxyEntries
            Write-AnalysisLog "  HttpProxy analysis complete: $($reportData.ServerResults[$serverName].HttpProxy.TotalRequests) requests" -Level Success
        } else {
            Write-AnalysisLog "  No HttpProxy entries collected for $serverName" -Level Warning
        }
    }
}

# Step 3: Generate reports
Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
    -Status "Phase 4/4: Generating HTML report..." -PercentComplete 85
Write-AnalysisLog "`n========== Generating Reports ==========" -Level Info

$htmlReport = New-IISHTMLReport -ReportData $reportData
$reportPath = Join-Path -Path $OutputPath -ChildPath "ExchangeIISAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlReport | Out-File -FilePath $reportPath -Encoding UTF8
Write-AnalysisLog "HTML report saved to: $reportPath" -Level Success

Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" `
    -Status "Phase 4/4: Exporting CSV data..." -PercentComplete 92

# CSV Export - IIS data
$csvData = @()
foreach ($serverName in $Servers) {
    $serverResult = $reportData.ServerResults[$serverName]
    
    if ($serverResult.IIS) {
        $iis = $serverResult.IIS
        foreach ($vdir in $iis.VDirStats.GetEnumerator()) {
            $csvData += [PSCustomObject]@{
                Timestamp       = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Server          = $serverName
                DataSource      = 'IIS'
                VirtualDirectory = $vdir.Key
                Requests        = $vdir.Value.Count
                AvgLatencyMs    = $vdir.Value.AvgTime
                P95LatencyMs    = $vdir.Value.P95Time
                Errors          = $vdir.Value.Errors
                ErrorRate       = $vdir.Value.ErrorRate
                SlowRequests    = $vdir.Value.SlowCount
            }
        }
    }
    
    if ($serverResult.HttpProxy) {
        $proxy = $serverResult.HttpProxy
        foreach ($pt in $proxy.ProxyTypeStats.GetEnumerator()) {
            $csvData += [PSCustomObject]@{
                Timestamp       = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                Server          = $serverName
                DataSource      = 'HttpProxy'
                VirtualDirectory = $pt.Key
                Requests        = $pt.Value.Count
                AvgLatencyMs    = $pt.Value.AvgTotalLatency
                P95LatencyMs    = $pt.Value.P95TotalLatency
                Errors          = $pt.Value.Errors
                ErrorRate       = $pt.Value.ErrorRate
                SlowRequests    = 'N/A'
            }
        }
    }
}

if ($csvData.Count -gt 0) {
    $csvPath = Join-Path -Path $OutputPath -ChildPath "ExchangeIISAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-AnalysisLog "CSV data exported to: $csvPath" -Level Success
}

Write-Progress -Id 0 -Activity "Exchange IIS Log Analysis" -Completed

# Summary
Write-AnalysisLog "`n========== Analysis Summary ==========" -Level Info
foreach ($serverName in $Servers) {
    $sr = $reportData.ServerResults[$serverName]
    if ($sr.IIS) {
        $errorPct = [math]::Round($sr.IIS.ErrorRequests / [math]::Max($sr.IIS.TotalRequests, 1) * 100, 1)
        Write-AnalysisLog "  $serverName IIS: $($sr.IIS.TotalRequests) requests, ${errorPct}% errors, P95=$($sr.IIS.P95ResponseTime)ms, $($sr.IIS.SlowRequests) slow" -Level $(if ($errorPct -gt 5) { 'Error' } elseif ($errorPct -gt 1) { 'Warning' } else { 'Success' })
    }
    if ($sr.HttpProxy) {
        Write-AnalysisLog "  $serverName HttpProxy: $($sr.HttpProxy.TotalRequests) requests, AvgBackend=$($sr.HttpProxy.AvgBackendLatency)ms, P95Total=$($sr.HttpProxy.P95TotalLatency)ms" -Level Info
    }
}

Write-AnalysisLog "`nAnalysis complete! Opening report..." -Level Success
Write-AnalysisLog "Report location: $reportPath" -Level Info

Invoke-Item $reportPath

#endregion
