# Quick Start Guide - Exchange Performance Diagnostics

## For Customer Calls - Immediate Use

### Step 1: Open Exchange Management Shell
```powershell
# On Exchange Server, open Exchange Management Shell (Run as Administrator)
# OR from PowerShell:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

### Step 2: Run the Diagnostic Tool
```powershell
# Navigate to script location
cd "C:\Github\Collaboration\Exchange On-Prem"

# Basic run (analyzes local server)
.\Get-ExchangePerformanceDiagnostics.ps1

# For DAG with two servers (recommended)
.\Get-ExchangePerformanceDiagnostics.ps1 -Servers "EXCH01", "EXCH02"
```

### Step 3: Review the HTML Report
- Report automatically opens in browser
- Look for RED (Critical) and YELLOW (Warning) indicators
- Check the summary at the top for issue counts

## What to Look For - Outlook Slowness

### Critical Metrics (in order of importance)

1. **RPC Latency** - Most common cause of Outlook slowness
   - Healthy: <50ms
   - Warning: 50-100ms
   - Critical: >100ms
   - **Action if high**: Check disk latency next

2. **Disk I/O Latency** - Often the root cause
   - Database Read/Write: <20ms healthy, >50ms critical
   - Log Write: <10ms healthy, >50ms critical
   - **Action if high**: Check storage performance, competing I/O

3. **CPU Usage** - Can cause general slowness
   - Healthy: <60%
   - Warning: 60-80%
   - Critical: >80%
   - **Action if high**: Check anti-virus, runaway processes

4. **Memory Pressure** - Causes performance degradation
   - Healthy: <80% committed
   - Warning: 80-90%
   - Critical: >90%
   - **Action if high**: Check w3wp memory, consider adding RAM

5. **DAG Replication Queues** - Can impact active database
   - Healthy: <10 logs
   - Warning: 10-50 logs
   - Critical: >50 logs
   - **Action if high**: Check network between servers, disk on passive

## Quick Diagnosis Flowchart

```
Users report Outlook slowness
    ↓
Run diagnostic tool
    ↓
Check RPC Latency
    ↓
├─ >100ms? → Check Disk Latency
│   ↓
│   ├─ >50ms? → STORAGE ISSUE (most common)
│   │   └─ Action: Check storage array, IOPS, competing I/O
│   │
│   └─ <20ms? → Check CPU/Memory
│       └─ Action: Review processes, anti-virus, memory usage
│
├─ 50-100ms? → Check all metrics
│   └─ Action: Look for combination of issues
│
└─ <50ms? → Check DAG replication and Content Index
    └─ Action: May be client-side or network issue
```

## Common Scenarios

### Scenario 1: High RPC + High Disk Latency
**Root Cause**: Storage bottleneck
**Immediate Actions**:
- Check storage array performance
- Review disk queue lengths
- Look for competing I/O (backups, anti-virus)
- Consider moving databases to faster storage

### Scenario 2: High RPC + High CPU
**Root Cause**: Processing bottleneck
**Immediate Actions**:
- Verify anti-virus exclusions
- Check for runaway processes
- Review search indexing activity
- Check for transport queue backlogs

### Scenario 3: High RPC + Normal Everything Else
**Root Cause**: Possible network or client issue
**Immediate Actions**:
- Check network latency to DCs
- Test Outlook connectivity: `Test-OutlookConnectivity`
- Review client-side issues
- Check MAPI/HTTP configuration

### Scenario 4: DAG Replication Issues
**Root Cause**: Replication bottleneck affecting active database
**Immediate Actions**:
- Check network between DAG members
- Verify passive copy disk performance
- Review replication network configuration
- Check for failed database copies

## Immediate Remediation Actions

### If Disk Latency is High
```powershell
# Check current disk queue
Get-Counter '\LogicalDisk(*)\Current Disk Queue Length'

# Check for competing processes
Get-Process | Sort-Object -Property CPU -Descending | Select-Object -First 10

# Review storage array (vendor-specific tools)
```

### If CPU is High
```powershell
# Find top CPU consumers
Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 Name, CPU, Id

# Check anti-virus exclusions
# Verify these paths are excluded:
# - Exchange installation path
# - Database files
# - Log files
# - Transport database
```

### If Memory is High
```powershell
# Check w3wp memory usage
Get-Process w3wp | Select-Object Id, WorkingSet64, PrivateMemorySize64

# Recycle application pools if needed
Restart-WebAppPool -Name "MSExchangeOWAAppPool"
Restart-WebAppPool -Name "MSExchangeMapiFrontEndAppPool"
```

### If DAG Queues are High
```powershell
# Check replication status
Get-MailboxDatabaseCopyStatus -Server EXCH01 | ft Name, Status, CopyQueueLength, ReplayQueueLength

# Suspend and resume replication if stuck
Suspend-MailboxDatabaseCopy -Identity "DB01\EXCH02"
Resume-MailboxDatabaseCopy -Identity "DB01\EXCH02"

# Reseed if necessary (last resort)
Update-MailboxDatabaseCopy -Identity "DB01\EXCH02" -DeleteExistingFiles
```

## Customer Communication

### While Running Diagnostic (1-2 minutes)
"I'm running a comprehensive diagnostic tool that will analyze both Exchange servers. This will collect performance metrics including RPC latency, disk I/O, CPU, memory, and database replication status. It will take about 1-2 minutes to complete."

### After Getting Results
**If Critical Issues Found**:
"The diagnostic has identified [X] critical issues. The primary cause of the slowness appears to be [RPC latency/disk performance/CPU bottleneck]. I can see that [specific metric] is at [value], which is well above the recommended threshold of [threshold]."

**If No Critical Issues**:
"The diagnostic shows all Exchange server metrics are within normal ranges. This suggests the issue may be client-side, network-related, or intermittent. I recommend we [run Test-OutlookConnectivity/check client network/monitor over time]."

### Recommended Next Steps
1. Share the HTML report with customer
2. Explain the top 3 issues found
3. Provide immediate remediation steps
4. Schedule follow-up if needed
5. Document findings in ticket

## Files Generated

- **HTML Report**: Visual report with all findings (auto-opens)
- **CSV Data**: Raw metrics for trending/analysis
- Both saved to script directory or specified `-OutputPath`

## Tips for Success

✅ **DO**:
- Run during business hours when users are experiencing slowness
- Analyze both DAG members
- Save reports for trending
- Share HTML report with customer

❌ **DON'T**:
- Run only on one server in a DAG
- Make changes without understanding root cause
- Ignore warnings (they often become critical)
- Skip documentation

## Emergency Contacts

- Microsoft Exchange Documentation: https://learn.microsoft.com/en-us/exchange/exchange-server
- Performance Recommendations: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/deployment-ref/performance-recommendations
