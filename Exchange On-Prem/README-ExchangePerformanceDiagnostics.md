# Exchange Server 2016 Performance Diagnostics Tool

## Overview

Comprehensive PowerShell diagnostic tool designed to identify performance bottlenecks in Exchange Server 2016 environments, particularly in DAG (Database Availability Group) configurations when users report slowness in Outlook classic.

## Features

### 1. Exchange Health Monitoring
- Server component health status
- Service health verification
- DAG replication status
- Database copy queue lengths
- Content index state

### 2. Performance Counter Analysis
- **RPC Performance**: Averaged latency, request counts, operations per second
- **CPU Utilization**: Processor time across all cores
- **Memory Usage**: Available memory, committed bytes percentage
- **Disk I/O**: Database read/write latency, log write latency
- **Database Operations**: I/O latency for database instances

### 3. Client Access Server Diagnostics
- IIS Application Pool status (MSExchangeOWAAppPool, MSExchangeMapiFrontEndAppPool, etc.)
- Worker process (w3wp.exe) resource consumption
- Active connection monitoring

### 4. Mailbox Database Metrics
- Database mount status
- Database size and available space
- Background maintenance status
- Transport queue analysis

### 5. Network & Active Directory Connectivity
- Domain Controller latency testing
- LDAP, Kerberos, DNS port connectivity
- Network latency between servers

### 6. Professional HTML Reporting
- Color-coded status indicators (Green/Yellow/Red)
- Performance metrics dashboard
- Actionable recommendations
- Threshold-based warnings
- CSV export for trending analysis

## Prerequisites

- **PowerShell Version**: 5.1 or higher
- **Exchange Management Shell**: Loaded or available
- **Permissions**: Administrator rights on Exchange servers
- **WinRM**: Enabled for remote data collection
- **Network Access**: Ability to query performance counters remotely

## Installation

1. Copy `Get-ExchangePerformanceDiagnostics.ps1` to your Exchange server or management workstation
2. Ensure Exchange Management Shell is available
3. Run from an elevated PowerShell session

## Usage

### Basic Usage

Analyze the local Exchange server:

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1
```

### Specify Multiple Servers

Analyze specific Exchange servers in your DAG:

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1 -Servers "EXCH01", "EXCH02"
```

### Custom Output Location

Save reports to a specific directory:

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1 -Servers "EXCH01", "EXCH02" -OutputPath "C:\ExchangeReports"
```

### Extended Collection Period

Collect performance data over 5 minutes for better averaging:

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1 -CollectionDuration 300
```

### Skip Network Tests

Skip network latency tests (faster execution):

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1 -SkipNetworkTests
```

### Complete Example

```powershell
.\Get-ExchangePerformanceDiagnostics.ps1 `
    -Servers "EXCH01", "EXCH02" `
    -OutputPath "C:\Reports\Exchange" `
    -CollectionDuration 120
```

## Performance Thresholds

The tool uses Microsoft best practice thresholds:

| Metric | Warning | Critical | Notes |
|--------|---------|----------|-------|
| RPC Latency | >50ms | >100ms | Average RPC latency for client operations |
| CPU Usage | >60% | >80% | Sustained CPU usage percentage |
| Memory Committed | >80% | >90% | Percentage of committed memory |
| Database Disk Read | >20ms | >50ms | Average database read latency |
| Database Disk Write | >20ms | >50ms | Average database write latency |
| Log Disk Write | >10ms | >50ms | Transaction log write latency |
| DAG Copy Queue | >10 logs | >50 logs | Replication queue length |
| DAG Replay Queue | >10 logs | >50 logs | Replay queue length |
| DC Network Latency | >50ms | N/A | Latency to domain controllers |

## Output Files

The script generates two files in the specified output directory:

1. **HTML Report**: `ExchangePerformanceReport_YYYYMMDD_HHMMSS.html`
   - Comprehensive visual report with all metrics
   - Color-coded status indicators
   - Actionable recommendations
   - Automatically opens in default browser

2. **CSV Data**: `ExchangePerformanceData_YYYYMMDD_HHMMSS.csv`
   - Raw performance counter data
   - Suitable for Excel analysis or trending
   - Can be imported into monitoring systems

## Interpreting Results

### Status Colors

- **Green (Healthy)**: Metric is within normal operating parameters
- **Yellow (Warning)**: Metric exceeds warning threshold, monitor closely
- **Red (Critical)**: Metric exceeds critical threshold, immediate action required

### Common Issues and Resolutions

#### High RPC Latency (>50ms)

**Possible Causes:**
- Disk I/O bottleneck
- Database fragmentation
- Insufficient memory
- Network issues

**Recommended Actions:**
- Check disk latency metrics
- Run database maintenance
- Verify anti-virus exclusions
- Review network connectivity

#### CPU Bottleneck (>60%)

**Possible Causes:**
- Anti-virus scanning Exchange files
- Runaway processes
- Insufficient hardware resources
- Search indexing activity

**Recommended Actions:**
- Verify anti-virus exclusions are configured
- Check for processes consuming excessive CPU
- Review search indexing status
- Consider hardware upgrade

#### Memory Pressure (>80%)

**Possible Causes:**
- Over-committed Exchange server
- Memory leaks in w3wp processes
- Insufficient physical memory
- Too many databases per server

**Recommended Actions:**
- Check w3wp process memory usage
- Recycle application pools
- Add physical memory
- Redistribute databases

#### Disk Latency (>20ms)

**Possible Causes:**
- Slow storage subsystem
- Competing I/O from other applications
- Storage array issues
- Insufficient IOPS

**Recommended Actions:**
- Verify storage performance
- Check for competing I/O
- Review storage array configuration
- Consider SSD upgrade for databases

#### DAG Replication Issues

**Possible Causes:**
- Network bandwidth limitations
- Disk I/O bottleneck on passive copy
- Network latency between DAG members
- Failed database copy

**Recommended Actions:**
- Check network connectivity between servers
- Verify disk performance on passive copies
- Review replication network configuration
- Reseed failed database copies if necessary

#### Content Index Not Healthy

**Possible Causes:**
- Indexing service issues
- Corrupted catalog
- Disk space issues
- Service crashes

**Recommended Actions:**
- Restart Microsoft Search service
- Rebuild search catalog
- Check disk space on catalog volume
- Review event logs for indexing errors

## Troubleshooting

### Script Fails to Load Exchange Cmdlets

**Error**: "The term 'Get-MailboxDatabase' is not recognized..."

**Solution**: 
```powershell
# Load Exchange Management Shell manually
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

Or run from Exchange Management Shell directly.

### Cannot Collect Performance Counters Remotely

**Error**: "Access is denied" or "RPC server unavailable"

**Solution**:
- Verify WinRM is enabled: `winrm quickconfig`
- Check firewall rules allow WinRM (TCP 5985/5986)
- Ensure you have administrator rights on target servers
- Verify Remote Registry service is running

### High Collection Duration Times Out

**Error**: Script takes too long or times out

**Solution**:
- Reduce `-CollectionDuration` parameter (default is 60 seconds)
- Use `-SkipNetworkTests` to speed up execution
- Run during off-peak hours

## Best Practices

1. **Regular Monitoring**: Run weekly during business hours to establish baseline
2. **Incident Response**: Run immediately when users report slowness
3. **Trending**: Keep CSV exports to track performance over time
4. **Documentation**: Save HTML reports for customer documentation
5. **Proactive**: Set up scheduled task to run daily and alert on critical issues

## Advanced Usage

### Scheduled Task Example

Create a scheduled task to run diagnostics daily:

```powershell
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' `
    -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Get-ExchangePerformanceDiagnostics.ps1" -Servers "EXCH01","EXCH02" -OutputPath "C:\Reports"'

$trigger = New-ScheduledTaskTrigger -Daily -At 9am

Register-ScheduledTask -Action $action -Trigger $trigger `
    -TaskName "Exchange Performance Diagnostics" `
    -Description "Daily Exchange performance analysis" `
    -User "DOMAIN\ServiceAccount" -RunLevel Highest
```

### Integration with Monitoring Systems

Export CSV data to your monitoring system:

```powershell
# Run diagnostic and import CSV
.\Get-ExchangePerformanceDiagnostics.ps1 -Servers "EXCH01","EXCH02"
$data = Import-Csv "C:\Path\To\ExchangePerformanceData_*.csv" | Sort-Object Timestamp -Descending | Select-Object -First 1

# Send to monitoring system (example)
if ($data.RPCLatency -gt 100) {
    Send-MailMessage -To "alerts@company.com" -Subject "CRITICAL: Exchange RPC Latency" -Body "RPC Latency: $($data.RPCLatency)ms"
}
```

## Support and Feedback

For issues or enhancements, refer to:
- [Microsoft Exchange Server Documentation](https://learn.microsoft.com/en-us/exchange/exchange-server)
- [Exchange Server Performance Recommendations](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/deployment-ref/performance-recommendations)
- [Exchange Server Health and Performance](https://learn.microsoft.com/en-us/exchange/high-availability/manage-ha/health-and-performance)

## Version History

- **v1.0** - Initial release
  - Exchange health monitoring
  - Performance counter collection
  - DAG replication analysis
  - HTML report generation
  - CSV export functionality

## License

This tool is provided as-is for Exchange Server administration and troubleshooting purposes.
