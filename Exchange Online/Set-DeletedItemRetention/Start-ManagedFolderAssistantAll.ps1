<#
.SYNOPSIS
    Runs Start-ManagedFolderAssistant -FullCrawl against all user mailboxes in Exchange Online.

.DESCRIPTION
    Enumerates all UserMailbox recipients via Get-ExoMailbox and triggers the Managed Folder
    Assistant (MRM processing) with a full crawl on each, using retry/backoff for throttling.
    Every result is logged to a CSV so interrupted runs can be resumed with -ResumeFromLog.

    Designed for large tenants (~21,000 mailboxes). Start-ManagedFolderAssistant is aggressively
    throttled in Exchange Online; expect roughly 1-2 operations per second, i.e. 3-6 hours for
    21k mailboxes. Use -ThrottleDelayMs to add a fixed pause between calls if you hit sustained
    throttling.

    Note: The command only QUEUES the mailbox for MFA processing. Actual retention/MRM processing
    happens asynchronously and can take hours to days to complete per mailbox.

.PARAMETER LogPath
    Path to the CSV result log. Defaults to a timestamped file next to the script.

.PARAMETER ResumeFromLog
    Path to a CSV log from a previous run. Mailboxes marked 'Started' in that log are skipped.

.PARAMETER MaxRetries
    Maximum retry attempts per mailbox on transient/throttling errors. Default 3.

.PARAMETER ThrottleDelayMs
    Optional fixed delay in milliseconds between mailboxes (0 = none). Default 0.
    Try 250-500 if you see continuous throttling warnings.

.PARAMETER MaxMailboxes
    Maximum number of mailboxes to trigger MFA on in this run (0 = no limit). Default 0.
    Skipped mailboxes (from -ResumeFromLog) do not count towards the limit, so you can
    process a large tenant in batches: run with -MaxMailboxes 5000, then resume with
    -ResumeFromLog and -MaxMailboxes 5000 again to do the next 5000.

.EXAMPLE
    .\Start-ManagedFolderAssistantAll.ps1 -WhatIf
    Dry run: shows which mailboxes would be processed without triggering MFA.

.EXAMPLE
    .\Start-ManagedFolderAssistantAll.ps1
    Triggers a full-crawl MFA run on all user mailboxes.

.EXAMPLE
    .\Start-ManagedFolderAssistantAll.ps1 -ResumeFromLog .\MFAFullCrawl_20260716_090000.csv
    Resumes a previous run, skipping mailboxes already triggered successfully.

.EXAMPLE
    .\Start-ManagedFolderAssistantAll.ps1 -MaxMailboxes 100
    Triggers MFA on the first 100 mailboxes only (useful as a pilot batch).

.NOTES
    Requires: ExchangeOnlineManagement module v3.4+ and Exchange Administrator (or Recipient Management) role.
    Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/start-managedfolderassistant
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$LogPath = (Join-Path $PSScriptRoot ("MFAFullCrawl_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [string]$ResumeFromLog,

    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,

    [ValidateRange(0, 10000)]
    [int]$ThrottleDelayMs = 0,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxMailboxes = 0
)

$ErrorActionPreference = 'Stop'

#region Helpers

function Write-Log {
    param(
        [Parameter(Mandatory)][hashtable]$Entry
    )
    [pscustomobject]$Entry | Export-Csv -Path $LogPath -Append -NoTypeInformation -Encoding UTF8
}

function Test-ExoConnection {
    try {
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
        Where-Object { $_.Status -eq 'Started' } |
        ForEach-Object { $processedGuids[$_.Guid] = $true }
    Write-Host ("Resume mode: {0} mailboxes already processed will be skipped." -f $processedGuids.Count) -ForegroundColor Yellow
}

#endregion

#region Enumerate mailboxes

Write-Host "Retrieving all user mailboxes (this can take several minutes in large tenants)..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$mailboxes = Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited `
    -Properties UserPrincipalName, ExchangeGuid

$total = ($mailboxes | Measure-Object).Count
Write-Host ("Found {0} user mailboxes in {1:mm\:ss}." -f $total, $sw.Elapsed) -ForegroundColor Cyan
Write-Host "Result log: $LogPath"

#endregion

#region Trigger Managed Folder Assistant

$counters = @{ Started = 0; Skipped = 0; Failed = 0 }
$index = 0
$limitReached = $false
if ($MaxMailboxes -gt 0) {
    Write-Host "Batch mode: will stop after triggering MFA on $MaxMailboxes mailboxes." -ForegroundColor Yellow
}
$sw.Restart()

foreach ($mbx in $mailboxes) {
    $index++
    $guid = $mbx.ExchangeGuid.ToString()
    $upn  = $mbx.UserPrincipalName

    # Progress + periodic summary
    if ($index % 25 -eq 0 -or $index -eq $total) {
        $pct = [math]::Round(($index / $total) * 100, 1)
        $rate = if ($sw.Elapsed.TotalSeconds -gt 0) { $index / $sw.Elapsed.TotalSeconds } else { 0 }
        $etaSeconds = if ($rate -gt 0) { ($total - $index) / $rate } else { 0 }
        Write-Progress -Activity "Starting Managed Folder Assistant (FullCrawl)" `
            -Status ("{0}/{1} ({2}%) - ETA {3:hh\:mm\:ss}" -f $index, $total, $pct, [timespan]::FromSeconds($etaSeconds)) `
            -PercentComplete $pct
    }
    if ($index % 500 -eq 0) {
        Write-Host ("[{0:HH:mm:ss}] {1}/{2} processed | Started: {3} | Skipped: {4} | Failed: {5}" -f `
            (Get-Date), $index, $total, $counters.Started, $counters.Skipped, $counters.Failed)
    }

    if ($processedGuids.ContainsKey($guid)) {
        $counters.Skipped++
        continue
    }

    if ($MaxMailboxes -gt 0 -and ($counters.Started + $counters.Failed) -ge $MaxMailboxes) {
        $limitReached = $true
        break
    }

    if (-not $PSCmdlet.ShouldProcess($upn, "Start-ManagedFolderAssistant -FullCrawl")) {
        continue
    }

    try {
        Invoke-WithRetry -Identity $upn -ScriptBlock {
            Start-ManagedFolderAssistant -Identity $guid -FullCrawl
        }
        $counters.Started++
        Write-Log @{
            Timestamp = (Get-Date -Format 'o')
            UPN       = $upn
            Guid      = $guid
            Status    = 'Started'
            Error     = ''
        }
    }
    catch {
        $counters.Failed++
        Write-Warning "FAILED: $upn - $($_.Exception.Message)"
        Write-Log @{
            Timestamp = (Get-Date -Format 'o')
            UPN       = $upn
            Guid      = $guid
            Status    = 'Failed'
            Error     = $_.Exception.Message
        }
    }

    if ($ThrottleDelayMs -gt 0) {
        Start-Sleep -Milliseconds $ThrottleDelayMs
    }
}

Write-Progress -Activity "Starting Managed Folder Assistant (FullCrawl)" -Completed

#endregion

#region Summary

Write-Host ""
Write-Host "================ Summary ================" -ForegroundColor Green
Write-Host ("Total mailboxes     : {0}" -f $total)
if ($limitReached) {
    Write-Host ("Batch limit reached : {0} (use -ResumeFromLog to continue)" -f $MaxMailboxes) -ForegroundColor Yellow
}
Write-Host ("Started             : {0}" -f $counters.Started)
Write-Host ("Skipped (resume)    : {0}" -f $counters.Skipped)
Write-Host ("Failed              : {0}" -f $counters.Failed) -ForegroundColor $(if ($counters.Failed -gt 0) { 'Red' } else { 'Green' })
Write-Host ("Elapsed             : {0:hh\:mm\:ss}" -f $sw.Elapsed)
Write-Host ("Log file            : {0}" -f $LogPath)
if ($counters.Failed -gt 0 -or $limitReached) {
    Write-Host ""
    Write-Host "Continue with remaining mailboxes (and retry failures) using:" -ForegroundColor Yellow
    Write-Host "  .\Start-ManagedFolderAssistantAll.ps1 -ResumeFromLog `"$LogPath`"" -ForegroundColor Yellow
}

#endregion
