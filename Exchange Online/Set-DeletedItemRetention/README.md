# Set-DeletedItemRetention

Sets `RetainDeletedItemsFor` (deleted item retention) to **30 days** on all user mailboxes in Exchange Online. Built for large tenants (~20,000 mailboxes): idempotent, resumable, and throttling-aware.

Reference: [Change deleted item retention for a mailbox](https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/change-deleted-item-retention)

## Prerequisites

- **PowerShell 7.x** recommended (Windows PowerShell 5.1 also works)
- **ExchangeOnlineManagement** module v3.4 or later:

  ```powershell
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  ```

- Admin role: **Exchange Administrator** or a role group containing Recipient Management (needs `Set-Mailbox`)
- Interactive sign-in (MFA supported) — the script calls `Connect-ExchangeOnline` if no session exists

> **Note:** 30 days is the maximum value for `RetainDeletedItemsFor` in Exchange Online.

## Usage

### 1. Dry run (recommended first)

```powershell
.\Set-DeletedItemRetention.ps1 -WhatIf
```

Shows which mailboxes would change without modifying anything.

### 2. Full run

```powershell
.\Set-DeletedItemRetention.ps1
```

- Enumerates all `UserMailbox` recipients with `Get-ExoMailbox -ResultSize Unlimited`
- Skips mailboxes already at ≥ 30 days (safe to re-run any time)
- Applies `Set-Mailbox -RetainDeletedItemsFor 30` with exponential-backoff retry on throttling
- Logs every result to a timestamped CSV next to the script

### 3. Resume an interrupted run / retry failures

```powershell
.\Set-DeletedItemRetention.ps1 -ResumeFromLog .\DeletedItemRetention_20260715_154700.csv
```

Mailboxes marked `Updated` or `AlreadyCompliant` in the log are skipped; everything else (including previous failures) is processed again.

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-RetentionDays` | `30` | Target retention (0–30 days) |
| `-LogPath` | Timestamped CSV next to script | Result log location |
| `-ResumeFromLog` | — | Previous run's CSV; successful entries are skipped |
| `-MaxRetries` | `3` | Retry attempts per mailbox on transient errors |
| `-WhatIf` | — | Dry run |

## Runtime expectations

`Set-Mailbox` runs serially at roughly 1–2 operations/second under Exchange Online throttling:

| Mailboxes | Estimated duration |
|---|---|
| 5,000 | ~1–1.5 hours |
| 20,000 | ~3–6 hours |

The script survives session/token expiry by reconnecting automatically, and interrupted runs can be resumed with `-ResumeFromLog`.

## Log format

CSV with columns: `Timestamp, UPN, Guid, PreviousValue, NewValue, Status, Error`.

`Status` values: `Updated`, `AlreadyCompliant`, `Failed`.

## Verify results

```powershell
Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties RetainDeletedItemsFor |
    Where-Object { $_.RetainDeletedItemsFor -lt (New-TimeSpan -Days 30) } |
    Select-Object UserPrincipalName, RetainDeletedItemsFor
```

An empty result means all user mailboxes are at 30 days.

## Important: new mailboxes

This script only changes **existing** mailboxes. New mailboxes will still get the **14-day default** from their mailbox plan. Either re-run this script periodically, or update the mailbox plans:

```powershell
Get-MailboxPlan | Set-MailboxPlan -RetainDeletedItemsFor 30
```

---

# Start-ManagedFolderAssistantAll

Companion script `Start-ManagedFolderAssistantAll.ps1` triggers `Start-ManagedFolderAssistant -FullCrawl` (MRM/retention processing) on all user mailboxes — useful after changing retention settings so policies are re-evaluated immediately instead of waiting for the normal MFA cycle (up to 7 days).

Reference: [Start-ManagedFolderAssistant](https://learn.microsoft.com/en-us/powershell/module/exchange/start-managedfolderassistant)

## Usage

```powershell
# Dry run
.\Start-ManagedFolderAssistantAll.ps1 -WhatIf

# Full run (all ~21,000 mailboxes)
.\Start-ManagedFolderAssistantAll.ps1

# Resume an interrupted run / retry failures
.\Start-ManagedFolderAssistantAll.ps1 -ResumeFromLog .\MFAFullCrawl_20260716_090000.csv

# Pilot batch: only the first 100 mailboxes
.\Start-ManagedFolderAssistantAll.ps1 -MaxMailboxes 100

# Process in batches of 5,000: run, then resume with the log to do the next 5,000
.\Start-ManagedFolderAssistantAll.ps1 -MaxMailboxes 5000
.\Start-ManagedFolderAssistantAll.ps1 -MaxMailboxes 5000 -ResumeFromLog .\MFAFullCrawl_20260716_090000.csv
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-LogPath` | Timestamped CSV next to script | Result log location |
| `-ResumeFromLog` | — | Previous run's CSV; entries marked `Started` are skipped |
| `-MaxRetries` | `3` | Retry attempts per mailbox on transient errors |
| `-ThrottleDelayMs` | `0` | Fixed pause between mailboxes (try 250–500 if throttled constantly) |
| `-MaxMailboxes` | `0` (no limit) | Stop after triggering MFA on this many mailboxes; combine with `-ResumeFromLog` for batching |
| `-WhatIf` | — | Dry run |

## Notes for large tenants

- Same resiliency pattern as the retention script: exponential-backoff retry, automatic reconnect on token expiry, resumable CSV log (`Status` values: `Started`, `Failed`).
- Expect roughly **1–2 operations/second** → **~3–6 hours for 21k mailboxes**.
- `Start-ManagedFolderAssistant` only **queues** the mailbox for processing — the actual MFA run is asynchronous and can take hours to days per mailbox to complete. There is no immediate "done" signal per mailbox.
- Check processing results per mailbox afterwards with:

  ```powershell
  Export-MailboxDiagnosticLogs -Identity user@contoso.com -ExtendedProperties |
      Select-Object -ExpandProperty MailboxLog |
      Select-String 'ELCLastSuccessTimestamp'
  ```
