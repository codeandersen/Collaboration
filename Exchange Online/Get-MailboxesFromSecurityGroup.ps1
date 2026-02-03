#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Sets retention policy on all mailboxes except those in the exempt security group.

.DESCRIPTION
    This script connects to Exchange Online using managed identity,
    processes mailboxes using streaming to avoid memory issues,
    excludes members of the exempt security group,
    and applies the specified retention policy in WhatIf mode.
    Optimized for large environments (25,000+ mailboxes).

.NOTES
    Designed for Azure Automation Account
    Requires ExchangeOnlineManagement module
    Runs in WhatIf mode by default
    Uses streaming to avoid memory issues
#>

param(
    [string]$ExemptSecurityGroup = "Exempt_OnlineArchive@starkworkspace.onmicrosoft.com",
    [string]$RetentionPolicyName = "STARK Group Default",
    [switch]$WhatIf = $true
)

try {
    Write-Output "========================================"
    Write-Output "Retention Policy Assignment Script"
    Write-Output "========================================"
    Write-Output "Exempt Security Group: $ExemptSecurityGroup"
    Write-Output "Retention Policy: $RetentionPolicyName"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "========================================`n"
    
    Write-Output "Connecting to Exchange Online using Managed Identity..."
    Connect-ExchangeOnline -ManagedIdentity -Organization "starkworkspace.onmicrosoft.com" -ShowBanner:$false
    Write-Output "Successfully connected to Exchange Online`n"
    
    Write-Output "Building exempt mailbox lookup table..."
    $exemptMailboxes = @{}
    
    try {
        $exemptMembers = Get-DistributionGroupMember -Identity $ExemptSecurityGroup -ResultSize Unlimited -ErrorAction Stop
        
        if ($exemptMembers) {
            $exemptCount = 0
            foreach ($member in $exemptMembers) {
                if ($member.PrimarySmtpAddress) {
                    $exemptMailboxes[$member.PrimarySmtpAddress.ToString().ToLowerInvariant()] = $true
                    $exemptCount++
                }
            }
            Write-Output "Found $exemptCount exempt mailbox(es)"
            $exemptMembers = $null
        }
        else {
            Write-Output "No exempt members found in the security group"
        }
    }
    catch {
        Write-Warning "Could not retrieve exempt members: $($_.Exception.Message)"
        Write-Output "Continuing without exemptions..."
    }
    
    [System.GC]::Collect()
    
    Write-Output "`nProcessing mailboxes using streaming (memory-efficient)..."
    Write-Output "========================================`n"
    
    $processedCount = 0
    $skippedCount = 0
    $assignedCount = 0
    $errorCount = 0
    $lastReportTime = Get-Date
    
    Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum | ForEach-Object {
        $processedCount++
        $emailAddress = $_.PrimarySmtpAddress.ToString().ToLowerInvariant()
        
        if ($exemptMailboxes.ContainsKey($emailAddress)) {
            $skippedCount++
            return
        }
        
        try {
            if ($WhatIf) {
                Set-Mailbox -Identity $_.Identity -RetentionPolicy $RetentionPolicyName -WhatIf
                $assignedCount++
            }
            else {
                Set-Mailbox -Identity $_.Identity -RetentionPolicy $RetentionPolicyName -ErrorAction Stop
                $assignedCount++
            }
        }
        catch {
            $errorCount++
            Write-Error "FAILED: $($_.DisplayName) ($($_.PrimarySmtpAddress)) - Error: $($_.Exception.Message)"
        }
        
        if ($processedCount % 500 -eq 0) {
            $currentTime = Get-Date
            $elapsed = ($currentTime - $lastReportTime).TotalSeconds
            $rate = if ($elapsed -gt 0) { [math]::Round(500 / $elapsed, 1) } else { 0 }
            Write-Output "Progress: $processedCount mailboxes processed | Assigned: $assignedCount | Skipped: $skippedCount | Errors: $errorCount | Rate: $rate/sec"
            $lastReportTime = $currentTime
            [System.GC]::Collect()
        }
    }
    
    Write-Output "`n========================================"
    Write-Output "Final Summary"
    Write-Output "========================================"
    Write-Output "Total mailboxes processed: $processedCount"
    Write-Output "Exempt mailboxes (skipped): $skippedCount"
    Write-Output "Mailboxes assigned policy: $assignedCount"
    Write-Output "Errors encountered: $errorCount"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "========================================`n"
    
    Write-Output "Disconnecting from Exchange Online..."
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Output "Successfully disconnected from Exchange Online"
    
    Write-Output "`nScript completed successfully"
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Could not disconnect from Exchange Online"
    }
    
    throw
}
