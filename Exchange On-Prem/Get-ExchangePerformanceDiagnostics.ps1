#Requires -Version 5.1
<#
.SYNOPSIS
    Comprehensive Exchange Server 2016 performance diagnostic tool for DAG environments.

.DESCRIPTION
    Collects and analyzes critical performance metrics from Exchange Server 2016 servers in a DAG configuration.
    Identifies bottlenecks causing Outlook slowness by examining:
    - Exchange health and DAG replication status
    - Performance counters (RPC latency, CPU, memory, disk I/O)
    - Client Access Server metrics
    - Mailbox database performance
    - Active Directory and network connectivity
    
    Generates an HTML report with color-coded warnings and actionable recommendations.

.PARAMETER Servers
    Array of Exchange server names to diagnose. Defaults to local server if not specified.

.PARAMETER OutputPath
    Directory path for output reports. Defaults to script directory.

.PARAMETER CollectionDuration
    Duration in seconds to collect performance counter samples for averaging. Default is 60 seconds.

.PARAMETER SkipNetworkTests
    Skip network latency and connectivity tests.

.EXAMPLE
    .\Get-ExchangePerformanceDiagnostics.ps1
    Analyzes the local Exchange server with default settings.

.EXAMPLE
    .\Get-ExchangePerformanceDiagnostics.ps1 -Servers "EXCH01","EXCH02" -OutputPath "C:\Reports"
    Analyzes two specific servers and saves reports to C:\Reports.

.EXAMPLE
    .\Get-ExchangePerformanceDiagnostics.ps1 -CollectionDuration 300
    Collects performance data over 5 minutes for better averaging.

.NOTES
    Author: Exchange Performance Diagnostics Tool
    Version: 1.0
    Requires: Exchange Management Shell or remote PowerShell session with Exchange cmdlets
    Permissions: Administrator rights on Exchange servers
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string[]]$Servers = @($env:COMPUTERNAME),
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = $PSScriptRoot,
    
    [Parameter(Mandatory = $false)]
    [int]$CollectionDuration = 60,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipNetworkTests
)

$ErrorActionPreference = 'Continue'
$WarningPreference = 'Continue'

$Script:DiagnosticResults = @{
    Timestamp = Get-Date
    Servers = @()
    Summary = @{
        CriticalIssues = 0
        Warnings = 0
        Healthy = 0
    }
}

#region Helper Functions

function Write-DiagnosticLog {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'Info' { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
    }
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Get-ThresholdStatus {
    param(
        [double]$Value,
        [double]$WarningThreshold,
        [double]$CriticalThreshold,
        [switch]$LowerIsBetter
    )
    
    if ($LowerIsBetter) {
        if ($Value -ge $CriticalThreshold) { return 'Critical' }
        elseif ($Value -ge $WarningThreshold) { return 'Warning' }
        else { return 'Healthy' }
    } else {
        if ($Value -le $CriticalThreshold) { return 'Critical' }
        elseif ($Value -le $WarningThreshold) { return 'Warning' }
        else { return 'Healthy' }
    }
}

#endregion

#region Exchange Health Functions

function Get-ExchangeHealthStatus {
    param([string]$ServerName)
    
    Write-DiagnosticLog "Collecting Exchange health status for $ServerName..." -Level Info
    
    $healthData = @{
        ServerName = $ServerName
        ComponentStates = @()
        ServiceHealth = @()
        ServerHealth = @()
        Status = 'Unknown'
    }
    
    try {
        $serverHealth = Get-ServerHealth -Identity $ServerName -ErrorAction Stop
        $healthData.ServerHealth = $serverHealth | Select-Object Server, HealthSet, AlertValue, State
        
        $serviceHealth = Test-ServiceHealth -Server $ServerName -ErrorAction Stop
        $healthData.ServiceHealth = $serviceHealth | Select-Object Role, RequiredServicesRunning, ServicesRunning, ServicesNotRunning
        
        $unhealthyComponents = $serverHealth | Where-Object { $_.AlertValue -ne 'Healthy' }
        if ($unhealthyComponents) {
            $healthData.Status = 'Warning'
            $Script:DiagnosticResults.Summary.Warnings++
            Write-DiagnosticLog "Found $($unhealthyComponents.Count) unhealthy components on $ServerName" -Level Warning
        } else {
            $healthData.Status = 'Healthy'
            $Script:DiagnosticResults.Summary.Healthy++
        }
        
    } catch {
        $healthData.Status = 'Error'
        $healthData.Error = $_.Exception.Message
        Write-DiagnosticLog "Error collecting health status: $($_.Exception.Message)" -Level Error
        $Script:DiagnosticResults.Summary.CriticalIssues++
    }
    
    return $healthData
}

function Get-DAGReplicationMetrics {
    param([string]$ServerName)
    
    Write-DiagnosticLog "Collecting DAG replication metrics for $ServerName..." -Level Info
    
    $dagData = @{
        ServerName = $ServerName
        DatabaseCopies = @()
        ReplicationIssues = @()
        Status = 'Unknown'
    }
    
    try {
        $dbCopies = Get-MailboxDatabaseCopyStatus -Server $ServerName -ErrorAction Stop
        
        foreach ($copy in $dbCopies) {
            $copyInfo = @{
                DatabaseName = $copy.DatabaseName
                Status = $copy.Status
                CopyQueueLength = $copy.CopyQueueLength
                ReplayQueueLength = $copy.ReplayQueueLength
                ContentIndexState = $copy.ContentIndexState
                ActivationPreference = $copy.ActivationPreference
            }
            
            $dagData.DatabaseCopies += $copyInfo
            
            if ($copy.CopyQueueLength -gt 50 -or $copy.ReplayQueueLength -gt 50) {
                $dagData.ReplicationIssues += "High queue length on $($copy.DatabaseName): Copy=$($copy.CopyQueueLength), Replay=$($copy.ReplayQueueLength)"
                $Script:DiagnosticResults.Summary.CriticalIssues++
            } elseif ($copy.CopyQueueLength -gt 10 -or $copy.ReplayQueueLength -gt 10) {
                $dagData.ReplicationIssues += "Elevated queue length on $($copy.DatabaseName): Copy=$($copy.CopyQueueLength), Replay=$($copy.ReplayQueueLength)"
                $Script:DiagnosticResults.Summary.Warnings++
            }
            
            if ($copy.ContentIndexState -ne 'Healthy') {
                $dagData.ReplicationIssues += "Content index not healthy on $($copy.DatabaseName): $($copy.ContentIndexState)"
                $Script:DiagnosticResults.Summary.Warnings++
            }
        }
        
        $dagData.Status = if ($dagData.ReplicationIssues.Count -eq 0) { 'Healthy' } else { 'Warning' }
        
    } catch {
        $dagData.Status = 'Error'
        $dagData.Error = $_.Exception.Message
        Write-DiagnosticLog "Error collecting DAG metrics: $($_.Exception.Message)" -Level Error
    }
    
    return $dagData
}

#endregion

#region Client Access Server Functions

function Get-CASMetrics {
    param([string]$ServerName)
    
    Write-DiagnosticLog "Collecting Client Access Server metrics for $ServerName..." -Level Info
    
    $casData = @{
        ServerName = $ServerName
        ApplicationPools = @()
        ActiveConnections = @{}
        Status = 'Unknown'
        Issues = @()
    }
    
    try {
        $appPools = Invoke-Command -ComputerName $ServerName -ScriptBlock {
            Import-Module WebAdministration -ErrorAction SilentlyContinue
            Get-ChildItem IIS:\AppPools | Where-Object { $_.Name -like '*Exchange*' } | Select-Object Name, State, @{N='WorkerProcesses';E={(Get-Process -Name w3wp -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'w3wp' }).Count}}
        } -ErrorAction Stop
        
        foreach ($pool in $appPools) {
            $poolInfo = @{
                Name = $pool.Name
                State = $pool.State
                WorkerProcesses = $pool.WorkerProcesses
            }
            $casData.ApplicationPools += $poolInfo
            
            if ($pool.State -ne 'Started') {
                $casData.Issues += "CRITICAL: Application pool $($pool.Name) is not started (State: $($pool.State))"
                $Script:DiagnosticResults.Summary.CriticalIssues++
            }
        }
        
        $w3wpProcesses = Invoke-Command -ComputerName $ServerName -ScriptBlock {
            Get-Process -Name w3wp -ErrorAction SilentlyContinue | Select-Object Id, CPU, WorkingSet64, Handles
        }
        
        foreach ($proc in $w3wpProcesses) {
            $memoryMB = [math]::Round($proc.WorkingSet64 / 1MB, 2)
            if ($memoryMB -gt 4096) {
                $casData.Issues += "WARNING: w3wp process (PID: $($proc.Id)) using $memoryMB MB of memory"
                $Script:DiagnosticResults.Summary.Warnings++
            }
        }
        
        $casData.Status = if ($casData.Issues.Count -eq 0) { 'Healthy' } else { 'Warning' }
        
    } catch {
        $casData.Status = 'Error'
        $casData.Error = $_.Exception.Message
        Write-DiagnosticLog "Error collecting CAS metrics: $($_.Exception.Message)" -Level Error
    }
    
    return $casData
}

#endregion

#region Mailbox Database Functions

function Get-MailboxDatabaseMetrics {
    param([string]$ServerName)
    
    Write-DiagnosticLog "Collecting mailbox database metrics for $ServerName..." -Level Info
    
    $dbData = @{
        ServerName = $ServerName
        Databases = @()
        Status = 'Unknown'
        Issues = @()
    }
    
    try {
        $databases = Get-MailboxDatabase -Server $ServerName -Status -ErrorAction Stop
        
        foreach ($db in $databases) {
            $dbInfo = @{
                Name = $db.Name
                Mounted = $db.Mounted
                DatabaseSize = [math]::Round($db.DatabaseSize.ToBytes() / 1GB, 2)
                AvailableNewMailboxSpace = if ($db.AvailableNewMailboxSpace) { [math]::Round($db.AvailableNewMailboxSpace.ToBytes() / 1GB, 2) } else { 0 }
                BackgroundDatabaseMaintenance = $db.BackgroundDatabaseMaintenance
            }
            $dbData.Databases += $dbInfo
            
            if (-not $db.Mounted) {
                $dbData.Issues += "CRITICAL: Database $($db.Name) is not mounted"
                $Script:DiagnosticResults.Summary.CriticalIssues++
            }
        }
        
        $queueData = Get-Queue -Server $ServerName -ErrorAction SilentlyContinue
        if ($queueData) {
            $largeQueues = $queueData | Where-Object { $_.MessageCount -gt 100 }
            foreach ($queue in $largeQueues) {
                $dbData.Issues += "WARNING: Queue $($queue.Identity) has $($queue.MessageCount) messages"
                $Script:DiagnosticResults.Summary.Warnings++
            }
        }
        
        $dbData.Status = if ($dbData.Issues.Count -eq 0) { 'Healthy' } else { 'Warning' }
        
    } catch {
        $dbData.Status = 'Error'
        $dbData.Error = $_.Exception.Message
        Write-DiagnosticLog "Error collecting database metrics: $($_.Exception.Message)" -Level Error
    }
    
    return $dbData
}

#endregion

#region Network and AD Functions

function Get-NetworkADMetrics {
    param([string]$ServerName)
    
    Write-DiagnosticLog "Collecting network and AD connectivity metrics for $ServerName..." -Level Info
    
    $netData = @{
        ServerName = $ServerName
        DomainControllers = @()
        NetworkLatency = @()
        Status = 'Unknown'
        Issues = @()
    }
    
    try {
        $dcs = Get-ADDomainController -Filter * -ErrorAction Stop | Select-Object -First 3
        
        foreach ($dc in $dcs) {
            $pingResult = Test-Connection -ComputerName $dc.HostName -Count 4 -ErrorAction SilentlyContinue
            if ($pingResult) {
                $avgLatency = ($pingResult | Measure-Object -Property ResponseTime -Average).Average
                $dcInfo = @{
                    Name = $dc.HostName
                    AverageLatency = [math]::Round($avgLatency, 2)
                    Status = if ($avgLatency -gt 50) { 'Warning' } else { 'Healthy' }
                }
                $netData.DomainControllers += $dcInfo
                
                if ($avgLatency -gt 50) {
                    $netData.Issues += "WARNING: High latency to DC $($dc.HostName): $([math]::Round($avgLatency, 2))ms"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
            } else {
                $netData.Issues += "CRITICAL: Cannot ping DC $($dc.HostName)"
                $Script:DiagnosticResults.Summary.CriticalIssues++
            }
        }
        
        $portTests = @(
            @{Port=389; Service='LDAP'},
            @{Port=88; Service='Kerberos'},
            @{Port=53; Service='DNS'}
        )
        
        foreach ($test in $portTests) {
            $result = Test-NetConnection -ComputerName $dcs[0].HostName -Port $test.Port -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            if (-not $result.TcpTestSucceeded) {
                $netData.Issues += "CRITICAL: Cannot connect to $($test.Service) (port $($test.Port)) on $($dcs[0].HostName)"
                $Script:DiagnosticResults.Summary.CriticalIssues++
            }
        }
        
        $netData.Status = if ($netData.Issues.Count -eq 0) { 'Healthy' } else { 'Warning' }
        
    } catch {
        $netData.Status = 'Error'
        $netData.Error = $_.Exception.Message
        Write-DiagnosticLog "Error collecting network/AD metrics: $($_.Exception.Message)" -Level Error
    }
    
    return $netData
}

#endregion

#region HTML Report Generation

function New-HTMLReport {
    param([hashtable]$DiagnosticData)
    
    Write-DiagnosticLog "Generating HTML report..." -Level Info
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $reportDate = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Performance Diagnostic Report - $reportDate</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #005a9e; margin-top: 30px; border-left: 5px solid #0078d4; padding-left: 10px; }
        h3 { color: #333; margin-top: 20px; }
        .summary { background-color: #fff; padding: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        .summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-top: 15px; }
        .summary-box { padding: 15px; border-radius: 5px; text-align: center; }
        .critical-box { background-color: #fde7e9; border: 2px solid #e81123; }
        .warning-box { background-color: #fff4ce; border: 2px solid #ffb900; }
        .healthy-box { background-color: #dff6dd; border: 2px solid: #107c10; }
        .summary-number { font-size: 36px; font-weight: bold; margin: 10px 0; }
        .summary-label { font-size: 14px; color: #666; }
        .server-section { background-color: #fff; padding: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; background-color: #fff; }
        th { background-color: #0078d4; color: white; padding: 12px; text-align: left; }
        td { padding: 10px; border-bottom: 1px solid #ddd; }
        tr:hover { background-color: #f5f5f5; }
        .status-healthy { color: #107c10; font-weight: bold; }
        .status-warning { color: #ffb900; font-weight: bold; }
        .status-critical { color: #e81123; font-weight: bold; }
        .issue-list { background-color: #fff; padding: 15px; border-left: 4px solid #e81123; margin: 10px 0; }
        .issue-critical { color: #e81123; font-weight: bold; }
        .issue-warning { color: #ffb900; font-weight: bold; }
        .recommendation { background-color: #e7f3ff; padding: 15px; border-left: 4px solid #0078d4; margin: 10px 0; }
        .metric-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin: 15px 0; }
        .metric-card { background-color: #f9f9f9; padding: 15px; border-radius: 5px; border: 1px solid #ddd; }
        .metric-label { font-size: 12px; color: #666; text-transform: uppercase; }
        .metric-value { font-size: 24px; font-weight: bold; margin: 5px 0; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 12px; text-align: center; }
    </style>
</head>
<body>
    <h1>Exchange Server 2016 Performance Diagnostic Report</h1>
    <div class="summary">
        <p><strong>Report Generated:</strong> $timestamp</p>
        <p><strong>Servers Analyzed:</strong> $($DiagnosticData.Servers.Count)</p>
        <div class="summary-grid">
            <div class="summary-box critical-box">
                <div class="summary-number">$($DiagnosticData.Summary.CriticalIssues)</div>
                <div class="summary-label">Critical Issues</div>
            </div>
            <div class="summary-box warning-box">
                <div class="summary-number">$($DiagnosticData.Summary.Warnings)</div>
                <div class="summary-label">Warnings</div>
            </div>
            <div class="summary-box healthy-box">
                <div class="summary-number">$($DiagnosticData.Summary.Healthy)</div>
                <div class="summary-label">Healthy Components</div>
            </div>
        </div>
    </div>
"@

    foreach ($server in $DiagnosticData.Servers) {
        $html += @"
    <div class="server-section">
        <h2>Server: $($server.ServerName)</h2>
"@

        if ($server.PerformanceCounters -and $server.PerformanceCounters.Status -eq 'Collected') {
            $perf = $server.PerformanceCounters
            $html += @"
        <h3>Performance Metrics</h3>
        <div class="metric-grid">
            <div class="metric-card">
                <div class="metric-label">RPC Averaged Latency</div>
                <div class="metric-value status-$($perf.RPC.Status.ToLower())">$(if ($perf.RPC.AverageLatency -eq 'N/A') { 'N/A' } else { "$($perf.RPC.AverageLatency) ms" })</div>
                <div class="metric-label">Threshold: 50ms (Warning), 100ms (Critical)</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">CPU Usage</div>
                <div class="metric-value status-$($perf.CPU.Status.ToLower())">$(if ($perf.CPU.AverageUsage -eq 'N/A') { 'N/A' } else { "$($perf.CPU.AverageUsage)%" })</div>
                <div class="metric-label">Threshold: 60% (Warning), 80% (Critical)</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Memory Committed</div>
                <div class="metric-value status-$($perf.Memory.Status.ToLower())">$(if ($perf.Memory.CommittedBytesPercent -eq 'N/A') { 'N/A' } else { "$($perf.Memory.CommittedBytesPercent)%" })</div>
                <div class="metric-label">Available: $(if ($perf.Memory.AvailableMB -eq 'N/A') { 'N/A' } else { "$($perf.Memory.AvailableMB) MB" })</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Disk Latency$(if ($perf.Disk.Source -eq 'LogicalDisk') { ' (LogicalDisk)' } elseif ($perf.Disk.Source -eq 'ExchangeDatabase') { ' (DB Counters)' })</div>
                <div class="metric-value status-$($perf.Disk.Status.ToLower())">R: $(if ($perf.Disk.DatabaseReadLatency -eq 'N/A') { 'N/A' } else { "$($perf.Disk.DatabaseReadLatency)ms" }) / W: $(if ($perf.Disk.DatabaseWriteLatency -eq 'N/A') { 'N/A' } else { "$($perf.Disk.DatabaseWriteLatency)ms" })</div>
                <div class="metric-label">Log Write: $(if ($perf.Disk.LogWriteLatency -eq 'N/A' -or ($perf.Disk.Source -eq 'LogicalDisk' -and $perf.Disk.LogWriteLatency -eq 0)) { 'N/A' } else { "$($perf.Disk.LogWriteLatency)ms" })</div>
            </div>
        </div>
"@
            if ($perf.Issues.Count -gt 0) {
                $html += "<div class='issue-list'><strong>Performance Issues:</strong><ul>"
                foreach ($issue in $perf.Issues) {
                    $class = if ($issue -like 'CRITICAL:*') { 'issue-critical' } else { 'issue-warning' }
                    $html += "<li class='$class'>$issue</li>"
                }
                $html += "</ul></div>"
            }
        }

        if ($server.DAGReplication) {
            $dag = $server.DAGReplication
            $html += @"
        <h3>DAG Replication Status</h3>
        <table>
            <tr>
                <th>Database</th>
                <th>Status</th>
                <th>Copy Queue</th>
                <th>Replay Queue</th>
                <th>Content Index</th>
                <th>Activation Preference</th>
            </tr>
"@
            foreach ($db in $dag.DatabaseCopies) {
                $statusClass = if ($db.Status -eq 'Healthy') { 'status-healthy' } elseif ($db.Status -eq 'Mounted') { 'status-healthy' } else { 'status-warning' }
                $html += @"
            <tr>
                <td>$($db.DatabaseName)</td>
                <td class="$statusClass">$($db.Status)</td>
                <td>$($db.CopyQueueLength)</td>
                <td>$($db.ReplayQueueLength)</td>
                <td>$($db.ContentIndexState)</td>
                <td>$($db.ActivationPreference)</td>
            </tr>
"@
            }
            $html += "</table>"
            
            if ($dag.ReplicationIssues.Count -gt 0) {
                $html += "<div class='issue-list'><strong>Replication Issues:</strong><ul>"
                foreach ($issue in $dag.ReplicationIssues) {
                    $html += "<li>$issue</li>"
                }
                $html += "</ul></div>"
            }
        }

        if ($server.MailboxDatabases) {
            $mdb = $server.MailboxDatabases
            $html += @"
        <h3>Mailbox Databases</h3>
        <table>
            <tr>
                <th>Database Name</th>
                <th>Mounted</th>
                <th>Size (GB)</th>
                <th>Available Space (GB)</th>
                <th>Background Maintenance</th>
            </tr>
"@
            foreach ($db in $mdb.Databases) {
                $mountedClass = if ($db.Mounted) { 'status-healthy' } else { 'status-critical' }
                $html += @"
            <tr>
                <td>$($db.Name)</td>
                <td class="$mountedClass">$($db.Mounted)</td>
                <td>$($db.DatabaseSize)</td>
                <td>$($db.AvailableNewMailboxSpace)</td>
                <td>$($db.BackgroundDatabaseMaintenance)</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        if ($server.CAS -and $server.CAS.ApplicationPools.Count -gt 0) {
            $html += @"
        <h3>Client Access Server - Application Pools</h3>
        <table>
            <tr>
                <th>Application Pool</th>
                <th>State</th>
                <th>Worker Processes</th>
            </tr>
"@
            foreach ($pool in $server.CAS.ApplicationPools) {
                $stateClass = if ($pool.State -eq 'Started') { 'status-healthy' } else { 'status-critical' }
                $html += @"
            <tr>
                <td>$($pool.Name)</td>
                <td class="$stateClass">$($pool.State)</td>
                <td>$($pool.WorkerProcesses)</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        if ($server.NetworkAD -and $server.NetworkAD.DomainControllers.Count -gt 0) {
            $html += @"
        <h3>Active Directory Connectivity</h3>
        <table>
            <tr>
                <th>Domain Controller</th>
                <th>Average Latency (ms)</th>
                <th>Status</th>
            </tr>
"@
            foreach ($dc in $server.NetworkAD.DomainControllers) {
                $statusClass = "status-$($dc.Status.ToLower())"
                $html += @"
            <tr>
                <td>$($dc.Name)</td>
                <td>$($dc.AverageLatency)</td>
                <td class="$statusClass">$($dc.Status)</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        if ($server.EventLogs -and $server.EventLogs.Status -eq 'Collected') {
            $evtLogs = $server.EventLogs
            $html += @"
        <h3>Exchange Event Logs (Last 24 Hours)</h3>
        <div class="metric-grid">
            <div class="metric-card">
                <div class="metric-label">Total Errors</div>
                <div class="metric-value status-critical">$($evtLogs.TotalErrors)</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Warnings</div>
                <div class="metric-value status-warning">$($evtLogs.TotalWarnings)</div>
            </div>
        </div>
"@
            if ($evtLogs.CriticalEvents.Count -gt 0) {
                $html += "<h4>Critical Events (Performance/Database Related)</h4><table><tr><th>Time</th><th>Source</th><th>Event ID</th><th>Message</th></tr>"
                foreach ($evt in $evtLogs.CriticalEvents | Select-Object -First 10) {
                    $html += @"
            <tr>
                <td>$($evt.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))</td>
                <td>$($evt.Source)</td>
                <td>$($evt.EventID)</td>
                <td>$($evt.Message)</td>
            </tr>
"@
                }
                $html += "</table>"
            }
            
            if ($evtLogs.Errors.Count -gt 0) {
                $html += "<h4>Recent Errors (Top 10)</h4><table><tr><th>Time</th><th>Source</th><th>Event ID</th><th>Message</th></tr>"
                foreach ($evt in $evtLogs.Errors | Select-Object -First 10) {
                    $html += @"
            <tr>
                <td>$($evt.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))</td>
                <td>$($evt.Source)</td>
                <td>$($evt.EventID)</td>
                <td>$($evt.Message)</td>
            </tr>
"@
                }
                $html += "</table>"
            }
        }

        $html += "</div>"
    }

    $html += @"
    <div class="server-section">
        <h2>Recommendations</h2>
        <div class="recommendation">
            <h3>Common Causes of Outlook Slowness in Exchange 2016:</h3>
            <ul>
                <li><strong>High RPC Latency (>50ms):</strong> Check disk I/O performance, database fragmentation, and consider database maintenance</li>
                <li><strong>CPU Bottleneck (>60%):</strong> Review anti-virus exclusions, check for runaway processes, consider hardware upgrade</li>
                <li><strong>Memory Pressure:</strong> Verify Exchange is not over-committed, check for memory leaks in w3wp processes</li>
                <li><strong>Disk Latency (>20ms):</strong> Verify storage performance, check for competing I/O, consider SSD upgrade</li>
                <li><strong>DAG Replication Issues:</strong> High queue lengths indicate network or disk bottlenecks between DAG members</li>
                <li><strong>Content Index Problems:</strong> Rebuild search indexes if not healthy</li>
                <li><strong>Network Latency to DCs:</strong> Ensure DCs are on same network segment, check for network congestion</li>
                <li><strong>Application Pool Issues:</strong> Recycle stuck pools, check IIS logs for errors</li>
            </ul>
        </div>
        <div class="recommendation">
            <h3>Immediate Actions Based on Findings:</h3>
            <ul>
"@

    if ($DiagnosticData.Summary.CriticalIssues -gt 0) {
        $html += "<li class='issue-critical'>Address $($DiagnosticData.Summary.CriticalIssues) critical issues immediately</li>"
    }
    if ($DiagnosticData.Summary.Warnings -gt 0) {
        $html += "<li class='issue-warning'>Review $($DiagnosticData.Summary.Warnings) warnings for potential problems</li>"
    }
    
    $html += @"
                <li>Run <code>Test-OutlookConnectivity</code> to verify client connectivity</li>
                <li>Check Event Viewer for Exchange-related errors</li>
                <li>Review IIS logs for failed requests or slow responses</li>
                <li>Monitor performance counters over time to identify trends</li>
                <li>Consider running Exchange Best Practices Analyzer</li>
            </ul>
        </div>
    </div>
    <div class="footer">
        <p>Exchange Performance Diagnostic Report | Generated: $timestamp</p>
        <p>For more information, visit <a href="https://learn.microsoft.com/en-us/exchange/exchange-server">Microsoft Exchange Documentation</a></p>
    </div>
</body>
</html>
"@

    return $html
}

#endregion

#region Main Execution

Write-DiagnosticLog "Starting Exchange Performance Diagnostics..." -Level Info
Write-DiagnosticLog "Servers to analyze: $($Servers -join ', ')" -Level Info
Write-DiagnosticLog "Collection duration: $CollectionDuration seconds" -Level Info

try {
    $snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
    if (-not $snapin) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        Write-DiagnosticLog "Loaded Exchange Management Shell snapin" -Level Success
    }
} catch {
    Write-DiagnosticLog "Warning: Could not load Exchange snapin. Assuming Exchange cmdlets are already available." -Level Warning
}

# Step 1: Start performance counter and event log collection for ALL servers as background jobs
# These are the slow operations (Get-Counter = 60+ sec, Get-WinEvent = several sec)
# Built-in cmdlets work fine in background jobs - no Exchange snapin needed
$perfJobs = @{}
$eventLogJobs = @{}

$counterGroups = @{
    'System' = @(
        '\Processor(_Total)\% Processor Time'
        '\Memory\Available MBytes'
        '\Memory\% Committed Bytes In Use'
    )
    'Disk' = @(
        '\LogicalDisk(*)\Avg. Disk sec/Read'
        '\LogicalDisk(*)\Avg. Disk sec/Write'
        '\LogicalDisk(*)\Current Disk Queue Length'
    )
    'ExchangeRPC' = @(
        '\MSExchange RpcClientAccess\RPC Averaged Latency'
        '\MSExchange RpcClientAccess\RPC Requests'
        '\MSExchange RpcClientAccess\RPC Operations/sec'
    )
}

foreach ($serverName in $Servers) {
    Write-DiagnosticLog "`n========== Starting background collection for $serverName ==========" -Level Info
    
    # Start perf counter jobs for each counter group on each server
    $perfJobs[$serverName] = @{}
    foreach ($groupName in $counterGroups.Keys) {
        $counters = $counterGroups[$groupName]
        $duration = $CollectionDuration
        $srv = $serverName
        
        $perfJobs[$serverName][$groupName] = Start-Job -ScriptBlock {
            param($ComputerName, $CounterList, $MaxSamples, $GrpName)
            try {
                $samples = Get-Counter -ComputerName $ComputerName -Counter $CounterList -SampleInterval 1 -MaxSamples $MaxSamples -ErrorAction Stop -WarningAction SilentlyContinue
                $validSamples = $samples.CounterSamples | Where-Object { $_.Status -eq 0 -or $null -eq $_.Status }
                return @{ Status = 'OK'; Samples = $validSamples; GroupName = $GrpName }
            } catch {
                return @{ Status = 'Error'; Message = $_.Exception.Message; GroupName = $GrpName }
            }
        } -ArgumentList $srv, $counters, $duration, $groupName
        
        Write-DiagnosticLog "  Started $groupName counter collection job for $serverName" -Level Info
    }
    
    # Start a separate ExchangeDatabase job that tries multiple counter path variants
    $perfJobs[$serverName]['ExchangeDatabase'] = Start-Job -ScriptBlock {
        param($ComputerName, $MaxSamples)
        
        # Try multiple counter path variants — different Exchange CU levels use different paths
        $counterVariants = @(
            @{
                Name = 'MSExchange Database ==> Instances (Attached)'
                Paths = @(
                    '\MSExchange Database ==> Instances(*)\I/O Database Reads (Attached) Average Latency'
                    '\MSExchange Database ==> Instances(*)\I/O Database Writes (Attached) Average Latency'
                    '\MSExchange Database ==> Instances(*)\I/O Log Writes Average Latency'
                )
            }
            @{
                Name = 'MSExchange Database ==> Instances'
                Paths = @(
                    '\MSExchange Database ==> Instances(*)\I/O Database Reads Average Latency'
                    '\MSExchange Database ==> Instances(*)\I/O Database Writes Average Latency'
                    '\MSExchange Database ==> Instances(*)\I/O Log Writes Average Latency'
                )
            }
            @{
                Name = 'MSExchangeIS Store'
                Paths = @(
                    '\MSExchangeIS Store(*)\RPC Latency average (msec)'
                    '\MSExchangeIS Store(*)\RPC Operations/sec'
                )
            }
            @{
                Name = 'Database (ESE)'
                Paths = @(
                    '\Database ==> Instances(*)\I/O Database Reads Average Latency'
                    '\Database ==> Instances(*)\I/O Database Writes Average Latency'
                    '\Database ==> Instances(*)\I/O Log Writes Average Latency'
                )
            }
        )
        
        foreach ($variant in $counterVariants) {
            try {
                $samples = Get-Counter -ComputerName $ComputerName -Counter $variant.Paths -SampleInterval 1 -MaxSamples $MaxSamples -ErrorAction Stop -WarningAction SilentlyContinue
                $validSamples = $samples.CounterSamples | Where-Object { $_.Status -eq 0 -or $null -eq $_.Status }
                if ($validSamples) {
                    return @{ Status = 'OK'; Samples = $validSamples; GroupName = 'ExchangeDatabase'; VariantUsed = $variant.Name }
                }
            } catch {
                # Try next variant
                continue
            }
        }
        
        return @{ Status = 'Error'; Message = 'No Exchange Database counter variants available on this server. Run lodctr /r on the Exchange server to rebuild performance counters.'; GroupName = 'ExchangeDatabase' }
    } -ArgumentList $srv, $CollectionDuration
    
    Write-DiagnosticLog "  Started ExchangeDatabase counter collection job (with fallbacks) for $serverName" -Level Info
    
    # Start event log collection job
    $eventLogJobs[$serverName] = Start-Job -ScriptBlock {
        param($ComputerName, $Hours)
        
        $exchangeLogs = @(
            'MSExchange ADAccess', 'MSExchangeIS', 'MSExchange Store Driver',
            'MSExchangeRepl', 'MSExchangeTransport', 'MSExchangeFrontEndTransport',
            'MSExchange ActiveSync', 'MSExchange OWA', 'MSExchangeAutodiscover',
            'MSExchangeMailboxReplication', 'MSExchangeThrottling'
        )
        
        $startTime = (Get-Date).AddHours(-$Hours)
        $errors = @()
        $warnings = @()
        
        foreach ($logName in $exchangeLogs) {
            try {
                $errEvents = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                    LogName = 'Application'; ProviderName = $logName; Level = 2; StartTime = $startTime
                } -MaxEvents 50 -ErrorAction SilentlyContinue
                
                if ($errEvents) {
                    foreach ($evt in $errEvents) {
                        $errors += [PSCustomObject]@{
                            TimeCreated = $evt.TimeCreated
                            Source = $evt.ProviderName
                            EventID = $evt.Id
                            Message = $evt.Message.Substring(0, [Math]::Min(200, $evt.Message.Length))
                        }
                    }
                }
                
                $warnEvents = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{
                    LogName = 'Application'; ProviderName = $logName; Level = 3; StartTime = $startTime
                } -MaxEvents 50 -ErrorAction SilentlyContinue
                
                if ($warnEvents) {
                    foreach ($evt in $warnEvents) {
                        $warnings += [PSCustomObject]@{
                            TimeCreated = $evt.TimeCreated
                            Source = $evt.ProviderName
                            EventID = $evt.Id
                            Message = $evt.Message.Substring(0, [Math]::Min(200, $evt.Message.Length))
                        }
                    }
                }
            } catch { }
        }
        
        return @{ Errors = $errors; Warnings = $warnings }
    } -ArgumentList $serverName, 24
    
    Write-DiagnosticLog "  Started event log collection job for $serverName" -Level Info
}

# Step 2: While background jobs run, collect Exchange cmdlet data sequentially (fast operations)
Write-DiagnosticLog "`n========== Collecting Exchange data while counters run in background ==========" -Level Info

$serverDataMap = @{}
foreach ($serverName in $Servers) {
    Write-DiagnosticLog "`n--- Collecting Exchange cmdlet data for $serverName ---" -Level Info
    
    $serverData = @{
        ServerName = $serverName
        Health = $null
        DAGReplication = $null
        PerformanceCounters = $null
        CAS = $null
        MailboxDatabases = $null
        NetworkAD = $null
        EventLogs = $null
    }
    
    $serverData.Health = Get-ExchangeHealthStatus -ServerName $serverName
    $serverData.DAGReplication = Get-DAGReplicationMetrics -ServerName $serverName
    $serverData.CAS = Get-CASMetrics -ServerName $serverName
    $serverData.MailboxDatabases = Get-MailboxDatabaseMetrics -ServerName $serverName
    
    if (-not $SkipNetworkTests) {
        $serverData.NetworkAD = Get-NetworkADMetrics -ServerName $serverName
    }
    
    $serverDataMap[$serverName] = $serverData
}

# Step 3: Wait for background jobs and merge results
Write-DiagnosticLog "`n========== Waiting for background collection jobs to complete ==========" -Level Info

foreach ($serverName in $Servers) {
    Write-DiagnosticLog "  Waiting for $serverName performance counter jobs..." -Level Info
    
    $perfData = @{
        ServerName = $serverName
        RPC = @{}
        CPU = @{}
        Memory = @{}
        Disk = @{}
        Network = @{}
        Status = 'Unknown'
        Issues = @()
    }
    
    $allSamples = @()
    
    foreach ($groupName in $perfJobs[$serverName].Keys) {
        $job = $perfJobs[$serverName][$groupName]
        $jobResult = Receive-Job -Job $job -Wait
        Remove-Job -Job $job -Force
        
        if ($jobResult.Status -eq 'OK' -and $jobResult.Samples) {
            $allSamples += $jobResult.Samples
            $variantMsg = if ($jobResult.VariantUsed) { " (using $($jobResult.VariantUsed))" } else { '' }
            Write-DiagnosticLog "  Received $($jobResult.Samples.Count) samples from $groupName on ${serverName}${variantMsg}" -Level Info
        } elseif ($jobResult.Status -eq 'Error') {
            if ($jobResult.Message -like '*lodctr*') {
                Write-DiagnosticLog "  $groupName counters not registered on $serverName - run 'lodctr /r' on the server to rebuild counters" -Level Warning
            } elseif ($jobResult.Message -notlike '*not valid*' -and $jobResult.Message -notlike '*not found*') {
                Write-DiagnosticLog "  Warning: Could not collect $groupName counters on ${serverName}: $($jobResult.Message)" -Level Warning
            } else {
                Write-DiagnosticLog "  $groupName counters not available on $serverName (counters may not be registered)" -Level Info
            }
        }
    }
    
    # Process collected samples into perfData
    if ($allSamples.Count -gt 0) {
        try {
            $rpcSamples = $allSamples | Where-Object { $_.Path -like '*RPC Averaged Latency*' -and $_.CookedValue -ge 0 }
            if ($rpcSamples) {
                $rpcLatency = ($rpcSamples | Measure-Object -Property CookedValue -Average).Average
                $rpcRequests = ($allSamples | Where-Object { $_.Path -like '*RPC Requests' } | Measure-Object -Property CookedValue -Average).Average
                $rpcOps = ($allSamples | Where-Object { $_.Path -like '*RPC Operations/sec*' } | Measure-Object -Property CookedValue -Average).Average
                
                $perfData.RPC = @{
                    AverageLatency = [math]::Round($rpcLatency, 2)
                    Requests = [math]::Round($rpcRequests, 0)
                    OperationsPerSec = [math]::Round($rpcOps, 2)
                    Status = Get-ThresholdStatus -Value $rpcLatency -WarningThreshold 50 -CriticalThreshold 100 -LowerIsBetter
                }
                
                if ($perfData.RPC.Status -eq 'Critical') {
                    $perfData.Issues += "CRITICAL: RPC latency is $($perfData.RPC.AverageLatency)ms (threshold: 100ms)"
                    $Script:DiagnosticResults.Summary.CriticalIssues++
                } elseif ($perfData.RPC.Status -eq 'Warning') {
                    $perfData.Issues += "WARNING: RPC latency is $($perfData.RPC.AverageLatency)ms (threshold: 50ms)"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
            } else {
                $perfData.RPC = @{ AverageLatency = 'N/A'; Requests = 'N/A'; OperationsPerSec = 'N/A'; Status = 'NotAvailable' }
            }
            
            $cpuSamples = $allSamples | Where-Object { $_.Path -like '*Processor(_Total)*' }
            if ($cpuSamples) {
                $cpuUsage = ($cpuSamples | Measure-Object -Property CookedValue -Average).Average
                $perfData.CPU = @{
                    AverageUsage = [math]::Round($cpuUsage, 2)
                    Status = Get-ThresholdStatus -Value $cpuUsage -WarningThreshold 60 -CriticalThreshold 80 -LowerIsBetter
                }
                
                if ($perfData.CPU.Status -eq 'Critical') {
                    $perfData.Issues += "CRITICAL: CPU usage is $($perfData.CPU.AverageUsage)% (threshold: 80%)"
                    $Script:DiagnosticResults.Summary.CriticalIssues++
                } elseif ($perfData.CPU.Status -eq 'Warning') {
                    $perfData.Issues += "WARNING: CPU usage is $($perfData.CPU.AverageUsage)% (threshold: 60%)"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
            } else {
                $perfData.CPU = @{ AverageUsage = 'N/A'; Status = 'NotAvailable' }
            }
            
            $memSamples = $allSamples | Where-Object { $_.Path -like '*Available MBytes*' }
            $commitSamples = $allSamples | Where-Object { $_.Path -like '*% Committed Bytes In Use*' }
            if ($memSamples -and $commitSamples) {
                $availableMemoryMB = ($memSamples | Measure-Object -Property CookedValue -Average).Average
                $committedBytes = ($commitSamples | Measure-Object -Property CookedValue -Average).Average
                
                $perfData.Memory = @{
                    AvailableMB = [math]::Round($availableMemoryMB, 0)
                    CommittedBytesPercent = [math]::Round($committedBytes, 2)
                    Status = Get-ThresholdStatus -Value $committedBytes -WarningThreshold 80 -CriticalThreshold 90 -LowerIsBetter
                }
                
                if ($perfData.Memory.Status -eq 'Critical') {
                    $perfData.Issues += "CRITICAL: Memory usage is $($perfData.Memory.CommittedBytesPercent)% (threshold: 90%)"
                    $Script:DiagnosticResults.Summary.CriticalIssues++
                } elseif ($perfData.Memory.Status -eq 'Warning') {
                    $perfData.Issues += "WARNING: Memory usage is $($perfData.Memory.CommittedBytesPercent)% (threshold: 80%)"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
            } else {
                $perfData.Memory = @{ AvailableMB = 'N/A'; CommittedBytesPercent = 'N/A'; Status = 'NotAvailable' }
            }
            
            # Try Exchange Database-specific counters first (values already in ms)
            $dbReadSamples = $allSamples | Where-Object { $_.Path -like '*I/O Database Reads*Latency*' -and $_.CookedValue -ge 0 }
            $dbWriteSamples = $allSamples | Where-Object { $_.Path -like '*I/O Database Writes*Latency*' -and $_.CookedValue -ge 0 }
            $logWriteSamples = $allSamples | Where-Object { $_.Path -like '*I/O Log Writes*Latency*' -and $_.CookedValue -ge 0 }
            
            $diskSource = $null
            $dbReadLatency = 0
            $dbWriteLatency = 0
            $logWriteLatency = 0
            
            if ($dbReadSamples -or $dbWriteSamples -or $logWriteSamples) {
                # Exchange Database counters available (values in ms)
                $diskSource = 'ExchangeDatabase'
                $dbReadLatency = if ($dbReadSamples) { ($dbReadSamples | Measure-Object -Property CookedValue -Average).Average } else { 0 }
                $dbWriteLatency = if ($dbWriteSamples) { ($dbWriteSamples | Measure-Object -Property CookedValue -Average).Average } else { 0 }
                $logWriteLatency = if ($logWriteSamples) { ($logWriteSamples | Measure-Object -Property CookedValue -Average).Average } else { 0 }
                Write-DiagnosticLog "  Using Exchange Database counters for disk latency" -Level Info
            } else {
                # Fallback: use LogicalDisk Avg. Disk sec/Read|Write counters (values in seconds, convert to ms)
                # Try per-drive first (exclude _Total), fall back to _Total if no per-drive found
                $ldReadSamples = $allSamples | Where-Object { $_.Path -like '*avg. disk sec/read*' -and $_.Path -notlike '*_total*' }
                $ldWriteSamples = $allSamples | Where-Object { $_.Path -like '*avg. disk sec/write*' -and $_.Path -notlike '*_total*' }
                
                # If no per-drive samples, try including _Total
                if (-not $ldReadSamples -and -not $ldWriteSamples) {
                    $ldReadSamples = $allSamples | Where-Object { $_.Path -like '*avg. disk sec/read*' }
                    $ldWriteSamples = $allSamples | Where-Object { $_.Path -like '*avg. disk sec/write*' }
                }
                
                if ($ldReadSamples -or $ldWriteSamples) {
                    $diskSource = 'LogicalDisk'
                    # Convert from seconds to milliseconds
                    $dbReadLatency = if ($ldReadSamples) { ($ldReadSamples | Measure-Object -Property CookedValue -Average).Average * 1000 } else { 0 }
                    $dbWriteLatency = if ($ldWriteSamples) { ($ldWriteSamples | Measure-Object -Property CookedValue -Average).Average * 1000 } else { 0 }
                    $logWriteLatency = 0
                    Write-DiagnosticLog "  Using LogicalDisk counters for disk latency (Exchange DB counters not available)" -Level Info
                } else {
                    Write-DiagnosticLog "  No disk latency samples found in $($allSamples.Count) total samples" -Level Warning
                }
            }
            
            if ($null -ne $diskSource) {
                $perfData.Disk = @{
                    DatabaseReadLatency = [math]::Round($dbReadLatency, 2)
                    DatabaseWriteLatency = [math]::Round($dbWriteLatency, 2)
                    LogWriteLatency = [math]::Round($logWriteLatency, 2)
                    Source = $diskSource
                    Status = 'Healthy'
                }
                
                if ($dbReadLatency -gt 50 -or $dbWriteLatency -gt 50) {
                    $perfData.Disk.Status = 'Critical'
                    $perfData.Issues += "CRITICAL: Disk latency exceeds 50ms (Read: $([math]::Round($dbReadLatency, 2))ms, Write: $([math]::Round($dbWriteLatency, 2))ms)"
                    $Script:DiagnosticResults.Summary.CriticalIssues++
                } elseif ($dbReadLatency -gt 20 -or $dbWriteLatency -gt 20) {
                    $perfData.Disk.Status = 'Warning'
                    $perfData.Issues += "WARNING: Disk latency exceeds 20ms (Read: $([math]::Round($dbReadLatency, 2))ms, Write: $([math]::Round($dbWriteLatency, 2))ms)"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
                
                if ($logWriteLatency -gt 50) {
                    $perfData.Disk.Status = 'Critical'
                    $perfData.Issues += "CRITICAL: Log disk write latency is $([math]::Round($logWriteLatency, 2))ms (threshold: 50ms)"
                    $Script:DiagnosticResults.Summary.CriticalIssues++
                } elseif ($logWriteLatency -gt 10) {
                    if ($perfData.Disk.Status -ne 'Critical') { $perfData.Disk.Status = 'Warning' }
                    $perfData.Issues += "WARNING: Log disk write latency is $([math]::Round($logWriteLatency, 2))ms (threshold: 10ms)"
                    $Script:DiagnosticResults.Summary.Warnings++
                }
            } else {
                $perfData.Disk = @{ DatabaseReadLatency = 'N/A'; DatabaseWriteLatency = 'N/A'; LogWriteLatency = 'N/A'; Source = 'None'; Status = 'NotAvailable' }
            }
            
            $perfData.Status = 'Collected'
        } catch {
            $perfData.Status = 'PartialError'
            $perfData.Error = $_.Exception.Message
            Write-DiagnosticLog "Error processing performance counters for ${serverName}: $($_.Exception.Message)" -Level Error
        }
    } else {
        $perfData.Status = 'Error'
        $perfData.Error = 'No valid performance counter samples collected'
        Write-DiagnosticLog "No valid performance counters collected for $serverName" -Level Warning
    }
    
    $serverDataMap[$serverName].PerformanceCounters = $perfData
    
    # Collect event log results
    Write-DiagnosticLog "  Waiting for $serverName event log job..." -Level Info
    $evtResult = Receive-Job -Job $eventLogJobs[$serverName] -Wait
    Remove-Job -Job $eventLogJobs[$serverName] -Force
    
    $eventData = @{
        ServerName = $serverName
        Errors = if ($evtResult.Errors) { $evtResult.Errors } else { @() }
        Warnings = if ($evtResult.Warnings) { $evtResult.Warnings } else { @() }
        CriticalEvents = @()
        Status = 'Collected'
        TotalErrors = if ($evtResult.Errors) { $evtResult.Errors.Count } else { 0 }
        TotalWarnings = if ($evtResult.Warnings) { $evtResult.Warnings.Count } else { 0 }
    }
    
    $criticalKeywords = @('crash', 'failed', 'corruption', 'database', 'timeout', 'deadlock', 'performance', 'slow')
    $eventData.CriticalEvents = $eventData.Errors | Where-Object {
        $msg = $_.Message.ToLower()
        $criticalKeywords | Where-Object { $msg -like "*$_*" }
    } | Select-Object -First 10
    
    if ($eventData.TotalErrors -gt 50) {
        $Script:DiagnosticResults.Summary.CriticalIssues++
        Write-DiagnosticLog "  Found $($eventData.TotalErrors) errors in Exchange event logs on $serverName" -Level Error
    } elseif ($eventData.TotalErrors -gt 10) {
        $Script:DiagnosticResults.Summary.Warnings++
        Write-DiagnosticLog "  Found $($eventData.TotalErrors) errors in Exchange event logs on $serverName" -Level Warning
    } else {
        Write-DiagnosticLog "  Found $($eventData.TotalErrors) errors, $($eventData.TotalWarnings) warnings in event logs on $serverName" -Level Info
    }
    
    $serverDataMap[$serverName].EventLogs = $eventData
    
    Write-DiagnosticLog "`n========== Completed Analysis: $serverName ==========" -Level Success
    $Script:DiagnosticResults.Servers += $serverDataMap[$serverName]
}

Write-DiagnosticLog "`n========== Generating Reports ==========" -Level Info

$htmlReport = New-HTMLReport -DiagnosticData $Script:DiagnosticResults
$reportPath = Join-Path -Path $OutputPath -ChildPath "ExchangePerformanceReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

Write-DiagnosticLog "HTML report saved to: $reportPath" -Level Success

$csvData = @()
foreach ($server in $Script:DiagnosticResults.Servers) {
    if ($server.PerformanceCounters -and $server.PerformanceCounters.Status -eq 'Collected') {
        $csvData += [PSCustomObject]@{
            Timestamp = $Script:DiagnosticResults.Timestamp
            ServerName = $server.ServerName
            RPCLatency = $server.PerformanceCounters.RPC.AverageLatency
            RPCStatus = $server.PerformanceCounters.RPC.Status
            CPUUsage = $server.PerformanceCounters.CPU.AverageUsage
            CPUStatus = $server.PerformanceCounters.CPU.Status
            MemoryCommitted = $server.PerformanceCounters.Memory.CommittedBytesPercent
            MemoryAvailableMB = $server.PerformanceCounters.Memory.AvailableMB
            MemoryStatus = $server.PerformanceCounters.Memory.Status
            DiskReadLatency = $server.PerformanceCounters.Disk.DatabaseReadLatency
            DiskWriteLatency = $server.PerformanceCounters.Disk.DatabaseWriteLatency
            LogWriteLatency = $server.PerformanceCounters.Disk.LogWriteLatency
            DiskStatus = $server.PerformanceCounters.Disk.Status
        }
    }
}

if ($csvData.Count -gt 0) {
    $csvPath = Join-Path -Path $OutputPath -ChildPath "ExchangePerformanceData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-DiagnosticLog "CSV data exported to: $csvPath" -Level Success
}

Write-DiagnosticLog "`n========== Diagnostic Summary ==========" -Level Info
Write-DiagnosticLog "Critical Issues: $($Script:DiagnosticResults.Summary.CriticalIssues)" -Level $(if ($Script:DiagnosticResults.Summary.CriticalIssues -gt 0) { 'Error' } else { 'Info' })
Write-DiagnosticLog "Warnings: $($Script:DiagnosticResults.Summary.Warnings)" -Level $(if ($Script:DiagnosticResults.Summary.Warnings -gt 0) { 'Warning' } else { 'Info' })
Write-DiagnosticLog "Healthy Components: $($Script:DiagnosticResults.Summary.Healthy)" -Level Success

Write-DiagnosticLog "`nDiagnostic complete! Open the HTML report for detailed analysis." -Level Success
Write-DiagnosticLog "Report location: $reportPath" -Level Info

if ($Script:DiagnosticResults.Summary.CriticalIssues -gt 0) {
    Write-DiagnosticLog "`nWARNING: Critical issues detected that require immediate attention!" -Level Error
}

Invoke-Item $reportPath

#endregion
