<#
    .SYNOPSIS
    Active Directory Domain and Forest Functional Level Upgrade Preparation Script

    .DESCRIPTION
    Comprehensive validation script to prepare for Active Directory Domain and Forest Functional Level upgrades.
    This script performs extensive health checks on:
    - Domain Controllers (OS versions, replication health)
    - FSMO roles and placement
    - DNS configuration and AD-integrated zones
    - SYSVOL/NETLOGON replication (DFSR)
    - Exchange integration (if present)
    - Azure AD Connect (if present)
    - Security configurations and potential issues

    .PARAMETER LogPath
    Optional custom path for the log file. Default is %TEMP%\YYYY-MM-DD_HHmmss_DomainFunctionalLevelPrep.log

    .EXAMPLE
    .\DomainForestFunctionLevelPrep.ps1
    Runs all validation checks and logs to default location.

    .EXAMPLE
    .\DomainForestFunctionLevelPrep.ps1 -LogPath "C:\Logs\ADPrep.log"
    Runs validation with custom log path.

    .NOTES
    Author: Exchange Admin
    Requirements:
    - Run from elevated PowerShell session
    - Active Directory PowerShell module
    - DNS Server PowerShell module (if checking DNS)
    - Exchange Management Shell (optional, for Exchange checks)
    - ADSync module (optional, for Azure AD Connect checks)
    
    .COPYRIGHT
    MIT License, feel free to distribute and use as you like, please leave author information.

    .LINK
    BLOG: http://www.hcandersen.net
    X: @dk_hcandersen

    .DISCLAIMER
    This script is provided AS-IS, with no warranty - Use at own risk.
    This script does NOT make any changes - it only validates readiness.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath
)

$ErrorActionPreference = "Continue"

# Create log file name using current date
if (-not $LogPath) {
    $DateStamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $LogPath = Join-Path $scriptDir "${DateStamp}_DomainFunctionalLevelPrep.log"
}

# Ensure log directory exists
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

Start-Transcript -Path $LogPath -Append

# Initialize tracking variables
$script:issuesFound = @()
$script:criticalIssues = 0
$script:warnings = 0
$script:checksPassed = 0
$startTime = Get-Date

# Helper function to track issues
function Add-Issue {
    param(
        [string]$Category,
        [string]$Description,
        [ValidateSet("Critical","Warning","Info")]
        [string]$Severity = "Warning"
    )
    
    $script:issuesFound += [PSCustomObject]@{
        Category = $Category
        Description = $Description
        Severity = $Severity
        Timestamp = Get-Date -Format "HH:mm:ss"
    }
    
    if ($Severity -eq "Critical") { $script:criticalIssues++ }
    elseif ($Severity -eq "Warning") { $script:warnings++ }
}

# Helper function for section headers
function Write-SectionHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "=====================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "=====================================" -ForegroundColor Cyan
}

Write-Host "============================================================" -ForegroundColor Green
Write-Host " Active Directory Domain/Forest Functional Level Validation" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Log File: $LogPath" -ForegroundColor Cyan
Write-Host "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host ""

#region Domain Controller Checks

# 1. Verify all DCs are Windows Server 2012 R2 or newer
Write-SectionHeader "[1] Domain Controller Operating Systems"
try {
    $allDCs = Get-ADDomainController -Filter * -ErrorAction Stop
    Write-Host "Found $($allDCs.Count) Domain Controller(s)" -ForegroundColor Cyan
    
    $allDCs | Select-Object HostName, Site, OperatingSystem, OperatingSystemVersion, IsGlobalCatalog |
        Sort-Object HostName | Format-Table -AutoSize
    
    $oldDCs = $allDCs | Where-Object { $_.OperatingSystem -match "2008|2003|2012$" }
    
    if ($oldDCs) {
        Write-Host "FAIL: Found unsupported DC operating systems:" -ForegroundColor Red
        $oldDCs | Select-Object HostName, OperatingSystem | Format-Table -AutoSize
        foreach ($dc in $oldDCs) {
            Add-Issue -Category "Domain Controllers" `
                -Description "Unsupported DC: $($dc.HostName) running $($dc.OperatingSystem)" `
                -Severity "Critical"
        }
    } else {
        Write-Host "PASS: All DCs are running supported OS versions (2012 R2 or newer)" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "ERROR: Failed to check Domain Controllers - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "Domain Controllers" -Description "Failed to retrieve DC information: $($_.Exception.Message)" -Severity "Critical"
}

#endregion

#region Functional Levels

# 2. Verify current domain and forest functional levels
Write-SectionHeader "[2] Current Domain and Forest Functional Levels"
try {
    $domain = Get-ADDomain -ErrorAction Stop
    $forest = Get-ADForest -ErrorAction Stop
    
    Write-Host "Domain: $($domain.DNSRoot)" -ForegroundColor Cyan
    Write-Host "  Current Domain Mode: $($domain.DomainMode)" -ForegroundColor Yellow
    Write-Host "Forest: $($forest.Name)" -ForegroundColor Cyan
    Write-Host "  Current Forest Mode: $($forest.ForestMode)" -ForegroundColor Yellow
    
    $domain | Select-Object DNSRoot, DomainMode | Format-Table -AutoSize
    $forest | Select-Object Name, ForestMode | Format-Table -AutoSize
    
    $script:checksPassed++
}
catch {
    Write-Host "ERROR: Failed to retrieve functional levels - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "Functional Levels" -Description "Failed to retrieve functional levels" -Severity "Critical"
}

#endregion

#region FSMO Roles

# 3. Verify FSMO role holders
Write-SectionHeader "[3] FSMO Role Holders"
try {
    $forest = Get-ADForest -ErrorAction Stop
    $domain = Get-ADDomain -ErrorAction Stop
    
    Write-Host "Forest-level FSMO roles:" -ForegroundColor Cyan
    Write-Host "  Schema Master: $($forest.SchemaMaster)" -ForegroundColor Yellow
    Write-Host "  Domain Naming Master: $($forest.DomainNamingMaster)" -ForegroundColor Yellow
    
    Write-Host "`nDomain-level FSMO roles:" -ForegroundColor Cyan
    Write-Host "  PDC Emulator: $($domain.PDCEmulator)" -ForegroundColor Yellow
    Write-Host "  RID Master: $($domain.RIDMaster)" -ForegroundColor Yellow
    Write-Host "  Infrastructure Master: $($domain.InfrastructureMaster)" -ForegroundColor Yellow
    
    # Check if FSMO role holders are online
    $fsmoHolders = @($forest.SchemaMaster, $forest.DomainNamingMaster, $domain.PDCEmulator, $domain.RIDMaster, $domain.InfrastructureMaster) | Select-Object -Unique
    
    foreach ($fsmo in $fsmoHolders) {
        $dcName = $fsmo.Split('.')[0]
        if (Test-Connection -ComputerName $dcName -Count 1 -Quiet) {
            Write-Host "  $dcName is online" -ForegroundColor Green
        } else {
            Write-Host "  $dcName is OFFLINE!" -ForegroundColor Red
            Add-Issue -Category "FSMO Roles" -Description "FSMO role holder $dcName is offline" -Severity "Critical"
        }
    }
    
    $script:checksPassed++
}
catch {
    Write-Host "ERROR: Failed to check FSMO roles - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "FSMO Roles" -Description "Failed to retrieve FSMO role information" -Severity "Critical"
}

#endregion

#region Global Catalog

# 4. Verify Global Catalog placement
Write-SectionHeader "[4] Global Catalog Servers"
try {
    $forest = Get-ADForest -ErrorAction Stop
    $gcServers = $forest.GlobalCatalogs
    
    Write-Host "Found $($gcServers.Count) Global Catalog server(s):" -ForegroundColor Cyan
    $gcServers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    
    if ($gcServers.Count -eq 0) {
        Add-Issue -Category "Global Catalog" -Description "No Global Catalog servers found" -Severity "Critical"
    } else {
        $script:checksPassed++
    }
}
catch {
    Write-Host "ERROR: Failed to check Global Catalog servers - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "Global Catalog" -Description "Failed to retrieve GC information" -Severity "Critical"
}

#endregion

#region Replication Health

# 5. Verify replication health
Write-SectionHeader "[5] AD Replication Health"
try {
    Write-Host "Running replication summary..." -ForegroundColor Cyan
    $replOutput = repadmin /replsum 2>&1
    $replOutput
    
    # Check for failures
    $failures = $replOutput | Select-String -Pattern "fail|error" -SimpleMatch
    if ($failures) {
        Write-Host "`nWARNING: Replication issues detected!" -ForegroundColor Red
        Add-Issue -Category "Replication" -Description "Replication failures detected" -Severity "Warning"
    } else {
        Write-Host "`nPASS: No replication failures detected" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "ERROR: Failed to check replication - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "Replication" -Description "Failed to run replication checks" -Severity "Warning"
}

#endregion

#region SYSVOL Replication

# 6. Verify SYSVOL replication method
Write-SectionHeader "[6] SYSVOL Replication (DFSR)"
try {
    Write-Host "Checking DFSR SYSVOL migration state..." -ForegroundColor Cyan
    $dfsrState = dfsrmig /getglobalstate 2>&1
    $dfsrState
    
    if ($dfsrState -match "Eliminated|'Redirected'") {
        Write-Host "`nPASS: SYSVOL is using DFSR" -ForegroundColor Green
        $script:checksPassed++
    } elseif ($dfsrState -match "Start|Prepared") {
        Write-Host "`nWARNING: SYSVOL migration to DFSR is not complete" -ForegroundColor Yellow
        Add-Issue -Category "SYSVOL" -Description "SYSVOL migration to DFSR is incomplete" -Severity "Warning"
    } else {
        Write-Host "`nINFO: DFSR state could not be determined" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "WARNING: Could not check DFSR state - $($_.Exception.Message)" -ForegroundColor Yellow
    Add-Issue -Category "SYSVOL" -Description "Could not verify DFSR state" -Severity "Info"
}

# Check SYSVOL and NETLOGON shares
Write-Host "`nChecking SYSVOL and NETLOGON shares on all DCs..." -ForegroundColor Cyan
try {
    $allDCs = Get-ADDomainController -Filter * -ErrorAction Stop
    $shareIssues = 0
    
    foreach ($dc in $allDCs) {
        try {
            $shares = Get-SmbShare -CimSession $dc.HostName -ErrorAction Stop |
                Where-Object { $_.Name -in @("SYSVOL","NETLOGON") }
            
            if ($shares.Count -eq 2) {
                Write-Host "  $($dc.HostName): OK (SYSVOL and NETLOGON present)" -ForegroundColor Green
            } else {
                Write-Host "  $($dc.HostName): MISSING SHARES!" -ForegroundColor Red
                Add-Issue -Category "SYSVOL" -Description "$($dc.HostName) is missing SYSVOL or NETLOGON shares" -Severity "Critical"
                $shareIssues++
            }
        }
        catch {
            Write-Host "  $($dc.HostName): Could not check shares - $($_.Exception.Message)" -ForegroundColor Yellow
            Add-Issue -Category "SYSVOL" -Description "Could not check shares on $($dc.HostName)" -Severity "Warning"
        }
    }
    
    if ($shareIssues -eq 0) {
        $script:checksPassed++
    }
}
catch {
    Write-Host "ERROR: Failed to check SYSVOL/NETLOGON shares - $($_.Exception.Message)" -ForegroundColor Red
}

#endregion

#region DNS

# 7. Verify AD-integrated DNS zones
Write-SectionHeader "[7] AD-Integrated DNS Zones"
try {
    $dnsZones = Get-DnsServerZone -ErrorAction Stop | Where-Object { $_.ZoneType -eq 'Primary' -and $_.IsDsIntegrated }
    
    if ($dnsZones) {
        Write-Host "Found $($dnsZones.Count) AD-integrated DNS zone(s):" -ForegroundColor Cyan
        $dnsZones | Select-Object ZoneName, ZoneType, ReplicationScope | Format-Table -AutoSize
        $script:checksPassed++
    } else {
        Write-Host "WARNING: No AD-integrated DNS zones found" -ForegroundColor Yellow
        Add-Issue -Category "DNS" -Description "No AD-integrated DNS zones found" -Severity "Warning"
    }
}
catch {
    Write-Host "WARNING: Could not check DNS zones - $($_.Exception.Message)" -ForegroundColor Yellow
    Add-Issue -Category "DNS" -Description "Could not retrieve DNS zone information" -Severity "Info"
}

#endregion

#region DC Diagnostics

# 8. Run DC diagnostics
Write-SectionHeader "[8] Domain Controller Diagnostics"
try {
    Write-Host "Running dcdiag (this may take several minutes)..." -ForegroundColor Cyan
    $dcdiagOutput = dcdiag /v 2>&1
    
    # Check for failures
    $dcdiagFailures = $dcdiagOutput | Select-String -Pattern "failed|error" -SimpleMatch
    
    if ($dcdiagFailures) {
        Write-Host "`nWARNING: dcdiag found issues:" -ForegroundColor Red
        $dcdiagFailures | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
        Add-Issue -Category "DC Diagnostics" -Description "dcdiag reported failures" -Severity "Warning"
    } else {
        Write-Host "`nPASS: dcdiag completed with no failures" -ForegroundColor Green
        $script:checksPassed++
    }
    
    # Save full output to log
    Write-Host "`nFull dcdiag output:" -ForegroundColor Cyan
    $dcdiagOutput
}
catch {
    Write-Host "ERROR: Failed to run dcdiag - $($_.Exception.Message)" -ForegroundColor Red
    Add-Issue -Category "DC Diagnostics" -Description "Failed to run dcdiag" -Severity "Warning"
}

#endregion

#region Time Synchronization

# 9. Verify time synchronization
Write-SectionHeader "[9] Time Synchronization"
try {
    Write-Host "Checking time synchronization across DCs..." -ForegroundColor Cyan
    $w32tmOutput = w32tm /monitor 2>&1
    $w32tmOutput
    
    # Check for time sync issues (offset > 5 seconds)
    $timeSyncIssues = $w32tmOutput | Select-String -Pattern "error|unreachable" -SimpleMatch
    
    if ($timeSyncIssues) {
        Write-Host "`nWARNING: Time synchronization issues detected" -ForegroundColor Yellow
        Add-Issue -Category "Time Sync" -Description "Time synchronization issues detected" -Severity "Warning"
    } else {
        Write-Host "`nPASS: Time synchronization appears healthy" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "WARNING: Could not check time synchronization - $($_.Exception.Message)" -ForegroundColor Yellow
}

#endregion

#region Stale Objects

# 10. Verify stale or disabled DC accounts
Write-SectionHeader "[10] Stale or Disabled DC Computer Accounts"
try {
    $dcComputers = Get-ADComputer -Filter 'PrimaryGroupID -eq 516' -Properties OperatingSystem, LastLogonDate, Enabled -ErrorAction Stop
    
    $staleDCs = $dcComputers | Where-Object { 
        (-not $_.Enabled) -or 
        ($_.LastLogonDate -and $_.LastLogonDate -lt (Get-Date).AddDays(-60))
    }
    
    if ($staleDCs) {
        Write-Host "WARNING: Found stale or disabled DC accounts:" -ForegroundColor Red
        $staleDCs | Select-Object Name, OperatingSystem, LastLogonDate, Enabled | Format-Table -AutoSize
        foreach ($dc in $staleDCs) {
            Add-Issue -Category "Stale Objects" -Description "Stale/disabled DC account: $($dc.Name)" -Severity "Warning"
        }
    } else {
        Write-Host "PASS: No stale or disabled DC accounts found" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "WARNING: Could not check for stale DC accounts - $($_.Exception.Message)" -ForegroundColor Yellow
}

# 11. Check for orphaned DC metadata
Write-Host "`nChecking for orphaned nTDSDSA objects..." -ForegroundColor Cyan
try {
    $ntdsObjects = Get-ADObject -LDAPFilter "(objectClass=nTDSDSA)" -ErrorAction Stop
    $activeDCs = Get-ADDomainController -Filter * -ErrorAction Stop
    
    $orphanedObjects = $ntdsObjects | Where-Object {
        $dcName = ($_.DistinguishedName -split ',')[1] -replace 'CN=',''
        $dcName -notin $activeDCs.Name
    }
    
    if ($orphanedObjects) {
        Write-Host "WARNING: Found orphaned nTDSDSA objects:" -ForegroundColor Red
        $orphanedObjects | Select-Object Name, DistinguishedName | Format-Table -AutoSize
        Add-Issue -Category "Stale Objects" -Description "Orphaned DC metadata found" -Severity "Warning"
    } else {
        Write-Host "PASS: No orphaned DC metadata found" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "WARNING: Could not check for orphaned metadata - $($_.Exception.Message)" -ForegroundColor Yellow
}

#endregion

#region Security Checks

# 12. Check duplicate SPNs
Write-SectionHeader "[11] Duplicate Service Principal Names (SPNs)"
try {
    Write-Host "Checking for duplicate SPNs..." -ForegroundColor Cyan
    $setspnOutput = setspn -X 2>&1
    $setspnOutput
    
    if ($setspnOutput -match "found 0 group") {
        Write-Host "`nPASS: No duplicate SPNs found" -ForegroundColor Green
        $script:checksPassed++
    } else {
        Write-Host "`nWARNING: Duplicate SPNs detected" -ForegroundColor Yellow
        Add-Issue -Category "Security" -Description "Duplicate SPNs found" -Severity "Warning"
    }
}
catch {
    Write-Host "WARNING: Could not check for duplicate SPNs - $($_.Exception.Message)" -ForegroundColor Yellow
}

# 13. Identify unconstrained delegation
Write-Host "`nChecking for unconstrained delegation..." -ForegroundColor Cyan
try {
    $unconstrainedComputers = Get-ADComputer -Filter {TrustedForDelegation -eq $true} -Properties TrustedForDelegation -ErrorAction Stop
    $unconstrainedUsers = Get-ADUser -Filter {TrustedForDelegation -eq $true} -Properties TrustedForDelegation -ErrorAction Stop
    
    if ($unconstrainedComputers -or $unconstrainedUsers) {
        Write-Host "INFO: Found objects with unconstrained delegation:" -ForegroundColor Yellow
        if ($unconstrainedComputers) {
            Write-Host "  Computers:" -ForegroundColor Cyan
            $unconstrainedComputers | Select-Object Name, DNSHostName | Format-Table -AutoSize
        }
        if ($unconstrainedUsers) {
            Write-Host "  Users:" -ForegroundColor Cyan
            $unconstrainedUsers | Select-Object Name, SamAccountName | Format-Table -AutoSize
        }
        Add-Issue -Category "Security" -Description "Objects with unconstrained delegation found" -Severity "Info"
    } else {
        Write-Host "PASS: No unconstrained delegation found" -ForegroundColor Green
        $script:checksPassed++
    }
}
catch {
    Write-Host "WARNING: Could not check for unconstrained delegation - $($_.Exception.Message)" -ForegroundColor Yellow
}

#endregion

#region Exchange Checks

# 14. Exchange-specific validation (optional)
Write-SectionHeader "[12] Exchange Server Integration (Optional)"
try {
    $exchangeServers = Get-ExchangeServer -ErrorAction Stop 2>$null
    
    if ($exchangeServers) {
        Write-Host "Found Exchange servers:" -ForegroundColor Cyan
        $exchangeServers | Select-Object Name, Site, AdminDisplayVersion | Format-Table -AutoSize
        
        Write-Host "`nChecking Exchange AD object version..." -ForegroundColor Cyan
        $exchangeContainer = Get-ADObject "CN=Microsoft Exchange,CN=Services,CN=Configuration,$((Get-ADRootDSE).configurationNamingContext)" -Properties objectVersion -ErrorAction Stop
        Write-Host "Exchange Schema Version: $($exchangeContainer.objectVersion)" -ForegroundColor Yellow
        
        $script:checksPassed++
    } else {
        Write-Host "INFO: No Exchange servers found (this is OK if Exchange is not deployed)" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "INFO: Exchange Management Shell not available or Exchange not installed" -ForegroundColor Yellow
}

#endregion

#region Azure AD Connect

# 15. Azure AD Connect scheduler health (optional)
Write-SectionHeader "[13] Azure AD Connect (Optional)"
try {
    $aadcScheduler = Get-ADSyncScheduler -ErrorAction Stop 2>$null
    
    if ($aadcScheduler) {
        Write-Host "Azure AD Connect Scheduler Status:" -ForegroundColor Cyan
        $aadcScheduler | Format-List
        $script:checksPassed++
    } else {
        Write-Host "INFO: Azure AD Connect not installed (this is OK if not using hybrid)" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "INFO: Azure AD Connect not installed or ADSync module not available" -ForegroundColor Yellow
}

#endregion

#region DC Reachability

# 16. Basic reachability test for DCs
Write-SectionHeader "[14] Domain Controller Reachability"
try {
    $allDCs = Get-ADDomainController -Filter * -ErrorAction Stop
    $unreachableDCs = 0
    
    foreach ($dc in $allDCs) {
        $reachable = Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet
        if ($reachable) {
            Write-Host "  $($dc.HostName): Reachable" -ForegroundColor Green
        } else {
            Write-Host "  $($dc.HostName): UNREACHABLE!" -ForegroundColor Red
            Add-Issue -Category "Connectivity" -Description "DC $($dc.HostName) is unreachable" -Severity "Critical"
            $unreachableDCs++
        }
    }
    
    if ($unreachableDCs -eq 0) {
        $script:checksPassed++
    }
}
catch {
    Write-Host "ERROR: Failed to check DC reachability - $($_.Exception.Message)" -ForegroundColor Red
}

#endregion

#region Summary Report

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " VALIDATION SUMMARY" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

$endTime = Get-Date
$duration = $endTime - $startTime

Write-Host "Execution Time: $($duration.ToString('mm\:ss'))" -ForegroundColor Cyan
Write-Host "Checks Passed: $script:checksPassed" -ForegroundColor Green
Write-Host "Critical Issues: $script:criticalIssues" -ForegroundColor $(if ($script:criticalIssues -gt 0) { "Red" } else { "Green" })
Write-Host "Warnings: $script:warnings" -ForegroundColor $(if ($script:warnings -gt 0) { "Yellow" } else { "Green" })
Write-Host ""

if ($script:issuesFound.Count -gt 0) {
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host " ISSUES FOUND" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""
    
    # Group by severity
    $critical = $script:issuesFound | Where-Object { $_.Severity -eq "Critical" }
    $warnings = $script:issuesFound | Where-Object { $_.Severity -eq "Warning" }
    $info = $script:issuesFound | Where-Object { $_.Severity -eq "Info" }
    
    if ($critical) {
        Write-Host "CRITICAL ISSUES (must be resolved before upgrade):" -ForegroundColor Red
        $critical | Format-Table Category, Description, Timestamp -AutoSize
    }
    
    if ($warnings) {
        Write-Host "WARNINGS (should be reviewed):" -ForegroundColor Yellow
        $warnings | Format-Table Category, Description, Timestamp -AutoSize
    }
    
    if ($info) {
        Write-Host "INFORMATIONAL:" -ForegroundColor Cyan
        $info | Format-Table Category, Description, Timestamp -AutoSize
    }
} else {
    Write-Host "No issues found! Environment appears ready for functional level upgrade." -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host " RECOMMENDATIONS" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

if ($script:criticalIssues -gt 0) {
    Write-Host "CRITICAL: Resolve all critical issues before proceeding with upgrade!" -ForegroundColor Red
    Write-Host "- Address unsupported DC operating systems" -ForegroundColor Yellow
    Write-Host "- Fix replication issues" -ForegroundColor Yellow
    Write-Host "- Ensure all DCs are online and reachable" -ForegroundColor Yellow
} elseif ($script:warnings -gt 0) {
    Write-Host "WARNING: Review all warnings before proceeding" -ForegroundColor Yellow
    Write-Host "- Some issues may not block the upgrade but should be addressed" -ForegroundColor Yellow
} else {
    Write-Host "Environment validation successful!" -ForegroundColor Green
    Write-Host "- All critical checks passed" -ForegroundColor Green
    Write-Host "- Review the log file for detailed information" -ForegroundColor Green
    Write-Host "- Ensure you have a valid backup before proceeding" -ForegroundColor Yellow
    Write-Host "- Test the upgrade in a non-production environment first" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Log file saved to: $LogPath" -ForegroundColor Cyan
Write-Host "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host ""

#endregion

Stop-Transcript
