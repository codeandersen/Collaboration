<#
.SYNOPSIS
    Sets RetainDeletedItemsFor (deleted item retention) to a target value on all user mailboxes in Exchange Online.

.DESCRIPTION
    Enumerates all UserMailbox recipients via Get-ExoMailbox, skips mailboxes already at or above the
    target retention, and applies Set-Mailbox -RetainDeletedItemsFor with retry/backoff for throttling.
    Every result is logged to a CSV so interrupted runs can be resumed with -ResumeFromLog.

    Designed for large tenants (tested pattern for ~20,000 mailboxes). Expect roughly 1-2 operations
    per second due to Exchange Online throttling, i.e. 3-6 hours for 20k mailboxes.

.PARAMETER RetentionDays
    Target deleted item retention in days. Default 30 (the Exchange Online maximum).

.PARAMETER LogPath
    Path to the CSV result log. Defaults to a timestamped file next to the script.

.PARAMETER ResumeFromLog
    Path to a CSV log from a previous run. Mailboxes marked 'Updated' or 'AlreadyCompliant'
    in that log are skipped.

.PARAMETER MaxRetries
    Maximum retry attempts per mailbox on transient/throttling errors. Default 3.

.EXAMPLE
    .\Set-DeletedItemRetention.ps1 -WhatIf
    Dry run: shows what would change without modifying any mailbox.

.EXAMPLE
    .\Set-DeletedItemRetention.ps1
    Sets 30-day retention on all user mailboxes not already compliant.

.EXAMPLE
    .\Set-DeletedItemRetention.ps1 -ResumeFromLog .\DeletedItemRetention_20260715_154700.csv
    Resumes a previous run, skipping mailboxes already processed successfully.

.NOTES
    Requires: ExchangeOnlineManagement module v3.4+ and Exchange Administrator (or Recipient Management) role.
    Reference: https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/change-deleted-item-retention
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateRange(0, 30)]
    [int]$RetentionDays = 30,

    [string]$LogPath = (Join-Path $PSScriptRoot ("DeletedItemRetention_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [string]$ResumeFromLog,

    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3
)

$ErrorActionPreference = 'Stop'
$targetRetention = New-TimeSpan -Days $RetentionDays

#region Helpers

function Write-Log {
    param(
        [Parameter(Mandatory)][hashtable]$Entry
    )
    [pscustomobject]$Entry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}

function Test-ExoConnection {
    try {
        $null = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' }
        return [bool](Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })
    }
    catch {
        return $false
    }
}

function Connect-Exo {
    Write-Host "Connecting to Exchange Online (interactive sign-in)..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory)][string]$Identity
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            & $ScriptBlock
            return
        }
        catch {
            $message = $_.Exception.Message

            # Reconnect if the session dropped (token expiry during long runs)
            if ($message -match 'not connected|session|token|unauthorized' -and -not (Test-ExoConnection)) {
                Write-Warning "Session lost while processing '$Identity'. Reconnecting..."
                Connect-Exo
                continue
            }

            $isTransient = $message -match 'throttl|busy|timeout|temporarily|transient|too many requests|503|429'
            if ($attempt -le $MaxRetries -and $isTransient) {
                $delay = [math]::Pow(2, $attempt) * 5   # 10s, 20s, 40s...
                Write-Warning "Transient error on '$Identity' (attempt $attempt/$MaxRetries): $message. Retrying in $delay s..."
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

#endregion

#region Prerequisites

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
    Sort-Object Version -Descending | Select-Object -First 1

if (-not $module) {
    Write-Error ("ExchangeOnlineManagement module not found. Install it with:`n" +
        "  Install-Module ExchangeOnlineManagement -Scope CurrentUser")
    return
}
if ($module.Version -lt [version]'3.0.0') {
    Write-Error ("ExchangeOnlineManagement v3.x or later is required (found $($module.Version)). Update with:`n" +
        "  Update-Module ExchangeOnlineManagement")
    return
}
Import-Module ExchangeOnlineManagement

if (-not (Test-ExoConnection)) {
    Connect-Exo
}
else {
    Write-Host "Reusing existing Exchange Online connection." -ForegroundColor Cyan
}

#endregion

#region Resume handling

$processedGuids = @{}
if ($ResumeFromLog) {
    if (-not (Test-Path $ResumeFromLog)) {
        Write-Error "Resume log not found: $ResumeFromLog"
        return
    }
    Import-Csv $ResumeFromLog |
        Where-Object { $_.Status -in @('Updated', 'AlreadyCompliant') } |
        ForEach-Object { $processedGuids[$_.Guid] = $true }
    Write-Host ("Resume mode: {0} mailboxes already processed will be skipped." -f $processedGuids.Count) -ForegroundColor Yellow
}

#endregion

#region Enumerate mailboxes

Write-Host "Retrieving all user mailboxes (this can take several minutes in large tenants)..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
    -Properties RetainDeletedItemsFor, UserPrincipalName, ExchangeGuid

$total = ($mailboxes | Measure-Object).Count
Write-Host ("Found {0} user mailboxes in {1:mm\:ss}." -f $total, $sw.Elapsed) -ForegroundColor Cyan
Write-Host "Result log: $LogPath"

#endregion

#region Apply

$counters = @{ Updated = 0; AlreadyCompliant = 0; Skipped = 0; Failed = 0 }
$index = 0
$sw.Restart()

foreach ($mbx in $mailboxes) {
    $index++
    $guid = $mbx.ExchangeGuid.ToString()
    $upn  = $mbx.UserPrincipalName
    $previous = $mbx.RetainDeletedItemsFor

    # Progress + periodic summary
    if ($index % 25 -eq 0 -or $index -eq $total) {
        $pct = [math]::Round(($index / $total) * 100, 1)
        $rate = if ($sw.Elapsed.TotalSeconds -gt 0) { $index / $sw.Elapsed.TotalSeconds } else { 0 }
        $etaSeconds = if ($rate -gt 0) { ($total - $index) / $rate } else { 0 }
        Write-Progress -Activity "Setting deleted item retention to $RetentionDays days" `
            -Status ("{0}/{1} ({2}%) - ETA {3:hh\:mm\:ss}" -f $index, $total, $pct, [timespan]::FromSeconds($etaSeconds)) `
            -PercentComplete $pct
    }
    if ($index % 500 -eq 0) {
        Write-Host ("[{0:HH:mm:ss}] {1}/{2} processed | Updated: {3} | AlreadyCompliant: {4} | Failed: {5}" -f `
            (Get-Date), $index, $total, $counters.Updated, $counters.AlreadyCompliant, $counters.Failed)
    }

    if ($processedGuids.ContainsKey($guid)) {
        $counters.Skipped++
        continue
    }

    if ($previous -ge $targetRetention) {
        $counters.AlreadyCompliant++
        Write-Log @{
            Timestamp     = (Get-Date -Format 'o')
            UPN           = $upn
            Guid          = $guid
            PreviousValue = $previous.ToString()
            NewValue      = $previous.ToString()
            Status        = 'AlreadyCompliant'
            Error         = ''
        }
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($upn, "Set RetainDeletedItemsFor to $RetentionDays days (currently $previous)")) {
        continue
    }

    try {
        Invoke-WithRetry -Identity $upn -ScriptBlock {
            Set-Mailbox -Identity $guid -RetainDeletedItemsFor $RetentionDays
        }
        $counters.Updated++
        Write-Log @{
            Timestamp     = (Get-Date -Format 'o')
            UPN           = $upn
            Guid          = $guid
            PreviousValue = $previous.ToString()
            NewValue      = $targetRetention.ToString()
            Status        = 'Updated'
            Error         = ''
        }
    }
    catch {
        $counters.Failed++
        Write-Warning "FAILED: $upn - $($_.Exception.Message)"
        Write-Log @{
            Timestamp     = (Get-Date -Format 'o')
            UPN           = $upn
            Guid          = $guid
            PreviousValue = $previous.ToString()
            NewValue      = ''
            Status        = 'Failed'
            Error         = $_.Exception.Message
        }
    }
}

Write-Progress -Activity "Setting deleted item retention to $RetentionDays days" -Completed

#endregion

#region Summary

Write-Host ""
Write-Host "================ Summary ================" -ForegroundColor Green
Write-Host ("Total mailboxes     : {0}" -f $total)
Write-Host ("Updated             : {0}" -f $counters.Updated)
Write-Host ("Already compliant   : {0}" -f $counters.AlreadyCompliant)
Write-Host ("Skipped (resume)    : {0}" -f $counters.Skipped)
Write-Host ("Failed              : {0}" -f $counters.Failed) -ForegroundColor $(if ($counters.Failed -gt 0) { 'Red' } else { 'Green' })
Write-Host ("Elapsed             : {0:hh\:mm\:ss}" -f $sw.Elapsed)
Write-Host ("Log file            : {0}" -f $LogPath)
if ($counters.Failed -gt 0) {
    Write-Host ""
    Write-Host "Re-run failed mailboxes with:" -ForegroundColor Yellow
    Write-Host "  .\Set-DeletedItemRetention.ps1 -ResumeFromLog `"$LogPath`"" -ForegroundColor Yellow
}

#endregion
