# ACS Email Delivery Dashboard

Azure Monitor Workbook and Azure Portal Dashboard for monitoring Azure Communication Services (ACS) email delivery, using the `ACSEmailStatusUpdateOperational` Log Analytics table.

## Overview

| Resource | File | Purpose |
|---|---|---|
| Azure Monitor Workbook | `workbook.json` | Interactive 3-tab analysis tool |
| Azure Portal Dashboard | `dashboard.json` | Pinnable at-a-glance KPI tiles |
| Deploy script | `deploy.ps1` | Deploys both via Az PowerShell module |

### No additional solutions or data connectors required
Both resources query the `ACSEmailStatusUpdateOperational` table which is natively populated by ACS when diagnostic settings are configured.

---

## Prerequisites

- **Az PowerShell module** (`Install-Module Az` if not installed)
- **Contributor** (or Owner) role on the target resource group
- **Log Analytics workspace** with `ACSEmailStatusUpdateOperational` data
  - Enable via: ACS resource → Diagnostic settings → Send to Log Analytics

---

## Deploy

```powershell
# Simple deployment (single subscription)
.\deploy.ps1 `
    -SubscriptionId    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ResourceGroupName "rg-monitoring" `
    -WorkspaceName     "law-acs-prod"

# Cross-subscription with tenant ID (multi-tenant scenarios)
.\deploy.ps1 `
    -TenantId          "tttttttt-tttt-tttt-tttt-tttttttttttt" `
    -SubscriptionId    "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ResourceGroupName "rg-monitoring" `
    -WorkspaceName     "law-acs-prod" `
    -WorkspaceSubscriptionId   "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -WorkspaceResourceGroupName "rg-acs"
```

> **Multi-tenant:** Add `-TenantId` if you have access to multiple Azure AD tenants.  
> **Cross-subscription:** If the Log Analytics workspace is in a different subscription than where you're deploying the workbook/dashboard, specify `WorkspaceSubscriptionId`.  
> `WorkspaceSubscriptionId` defaults to `SubscriptionId` (same subscription).  
> `WorkspaceResourceGroupName` defaults to `ResourceGroupName` if omitted (same RG).

The script outputs direct portal URLs for both resources when complete.

---

## Workbook — Tab Reference

### Tab 1: Overview
- **Parameters:** Time Range (default 30d), DeliveryStatus multi-select
- **Delivery summary table** — per SenderDomain / SenderUsername with:
  - Counts: Total, Delivered, Bounced, Suppressed, Failed, Hard Bounces, OutForDelivery
  - Rates: DeliveryRate%, BounceRate%, SuppressionRate%, FailureRate%
  - Color-coded heat bars (green = good, red = bad)
- **Status distribution** — pie chart
- **Events over time** — stacked bar chart (1h buckets)

### Tab 2: Sender Search
- **Parameters:** Sender Domain (text, partial match), Sender Username (text, partial match)
- Filtered delivery stats table for matching senders
- Daily trend line chart for filtered senders

### Tab 3: Recipient / Message Lookup
- **Parameters:** Recipient Email (text, partial match), CorrelationId (text, partial match)
- **Latest status table** — one row per message (latest status), with color-coded DeliveryStatus and bounce type
- **Full history table** — all events chronologically (useful for tracing status progression)

---

## Dashboard — Tile Reference

| Tile | Time Window | Description |
|---|---|---|
| Total Emails | 24h | Unique recipients contacted |
| Delivered | 24h | Successfully delivered count |
| Bounced | 24h | Bounced count |
| Failed + Bounced | 24h | Combined problem count |
| Delivery Rate by Domain | 7d | Top 10 domains, bar chart |
| Status Trend | 7d | Line chart per DeliveryStatus (6h buckets) |

---

## Re-deploy / Update

Re-running `deploy.ps1` creates new resource GUIDs each time (new workbook/dashboard instances).  
To **update an existing** workbook in-place, edit it in the Azure Portal and use **Save** — or supply the existing GUID via `-workbookId` / `-dashboardId` parameters (add them to `deploy.ps1` if needed).

---

## Related KQL — Standalone Summary Query

```kql
let LatestStatus =
    ACSEmailStatusUpdateOperational
    | where TimeGenerated > ago(30d)
    | where OperationName == "DeliveryStatusUpdate"
    | where isnotempty(RecipientId)
    | summarize arg_max(TimeGenerated, *) by CorrelationId, RecipientId;
LatestStatus
| extend SubscriptionId = tostring(split(_ResourceId, "/")[2])
| summarize
    TotalRecipients = count(),
    Delivered       = countif(DeliveryStatus == "Delivered"),
    OutForDelivery  = countif(DeliveryStatus == "OutForDelivery"),
    Bounced         = countif(DeliveryStatus == "Bounced"),
    Suppressed      = countif(DeliveryStatus == "Suppressed"),
    Failed          = countif(DeliveryStatus == "Failed"),
    HardBounces     = countif(tobool(IsHardBounce) == true)
    by SubscriptionId, SenderDomain, SenderUsername, _ResourceId
| extend
    DeliveryRatePct    = round(100.0 * Delivered / TotalRecipients, 2),
    BounceRatePct      = round(100.0 * Bounced / TotalRecipients, 2),
    SuppressionRatePct = round(100.0 * Suppressed / TotalRecipients, 2),
    FailureRatePct     = round(100.0 * Failed / TotalRecipients, 2)
| order by SubscriptionId asc, TotalRecipients desc
```
