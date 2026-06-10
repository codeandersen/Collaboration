# Exchange IIS & HttpProxy Log Analyzer

Analyzes IIS and Exchange HttpProxy logs from Exchange 2016 servers to identify slow requests, HTTP errors, authentication failures, and top talkers contributing to Outlook performance issues.

## Why Use This Tool

All Exchange 2016 client traffic flows through IIS:

| Virtual Directory | Protocol | Client |
|---|---|---|
| `/mapi/` | MAPI over HTTP | Outlook 2013+ |
| `/rpc/` | RPC over HTTP | Outlook Anywhere |
| `/owa/` | HTTP | Outlook Web App |
| `/ews/` | SOAP/HTTP | Exchange Web Services |
| `/Microsoft-Server-ActiveSync/` | HTTP | Mobile Devices |
| `/autodiscover/` | HTTP | Client Auto-Configuration |

IIS logs show request volume, response times, and HTTP errors. Exchange HttpProxy logs add **backend latency**, **authentication latency**, and **proxy-layer errors** that IIS logs don't capture.

## Requirements

- **PowerShell 5.1+**
- **Network access** to Exchange server admin shares (`\\ServerName\C$`)
- **No external dependencies** (no LogParser required)

## Quick Start

```powershell
# Basic: Analyze IIS logs from last 24 hours
.\Analyze-ExchangeIISLogs.ps1 -Servers "EX01","EX02"

# With HttpProxy logs for deeper backend analysis
.\Analyze-ExchangeIISLogs.ps1 -Servers "EX01","EX02" -IncludeHttpProxy

# Last 4 hours, lower slow threshold
.\Analyze-ExchangeIISLogs.ps1 -Servers "EX01" -Hours 4 -SlowRequestThreshold 2000

# HttpProxy only (skip IIS logs)
.\Analyze-ExchangeIISLogs.ps1 -Servers "EX01" -IncludeHttpProxy -SkipIISLogs

# Analyze only Frontend IIS site
.\Analyze-ExchangeIISLogs.ps1 -Servers "EX01" -LogSite Frontend
```

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Servers` | string[] | **Mandatory** | Exchange server names |
| `-Hours` | int | 24 | Hours back to analyze |
| `-OutputPath` | string | Script directory | Where to save reports |
| `-SlowRequestThreshold` | int | 5000 | Slow request threshold in ms |
| `-LogSite` | string | Both | `Frontend` (W3SVC1), `Backend` (W3SVC2), or `Both` |
| `-IncludeHttpProxy` | switch | Off | Also analyze Exchange HttpProxy logs |
| `-SkipIISLogs` | switch | Off | Skip IIS logs (use with `-IncludeHttpProxy`) |

## What It Analyzes

### IIS Log Analysis
- **Request volume & latency** — Total requests, average/P95/max response time
- **Slow requests** — Requests exceeding the threshold, broken down by virtual directory
- **HTTP status codes** — Distribution with error rates
- **Virtual directory breakdown** — Per-vdir request count, latency, and error rate
- **Top talkers** — Top 20 users by request count and by average latency
- **Top client IPs** — Top 20 source IPs by request volume
- **Authentication failures** — 401 errors with user and client details
- **Server errors** — 500/502/503/504 errors with context

### HttpProxy Log Analysis (with `-IncludeHttpProxy`)
- **Backend latency** — Time spent on backend vs frontend processing
- **Authentication latency** — AuthModule processing time
- **Proxy type breakdown** — Per-protocol stats (Mapi, Ews, OwaProxy, RpcHttp, etc.)
- **Backend server distribution** — Which backend servers handle load and their latency
- **Proxy errors** — ErrorCode and GenericErrors from the proxy layer

## Output Files

| File | Description |
|---|---|
| `ExchangeIISAnalysis_YYYYMMDD_HHmmss.html` | Color-coded HTML dashboard |
| `ExchangeIISAnalysis_YYYYMMDD_HHmmss.csv` | Per-vdir summary data for trending |

## Log File Locations

### IIS Logs (W3C format)
- **Frontend (W3SVC1):** `C:\inetpub\logs\LogFiles\W3SVC1\`
- **Backend (W3SVC2):** `C:\inetpub\logs\LogFiles\W3SVC2\`

### HttpProxy Logs (CSV format)
Base path: `C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\`

| Subdirectory | Protocol |
|---|---|
| `Mapi` | MAPI over HTTP |
| `Ews` | Exchange Web Services |
| `OwaProxy` | Outlook Web App |
| `RpcHttp` | RPC over HTTP (Outlook Anywhere) |
| `Autodiscover` | Autodiscover |
| `ActiveSync` | ActiveSync |

## Interpreting Results

### Key Metrics to Watch

| Metric | Healthy | Warning | Critical |
|---|---|---|---|
| **P95 Response Time** | < 1000ms | 1000-5000ms | > 5000ms |
| **Error Rate** | < 1% | 1-5% | > 5% |
| **401 Errors** | Occasional | Repeated for same user | Widespread |
| **503 Errors** | None | Occasional | Frequent (app pool crash) |
| **Backend Latency** | < 500ms | 500-2000ms | > 2000ms |
| **Auth Latency** | < 100ms | 100-500ms | > 500ms |

### Common Issues

- **High `/mapi/` latency** — Outlook is slow; check backend Exchange server and database latency
- **Many 401 on `/autodiscover/`** — Kerberos/NTLM issues; check SPNs and auth configuration
- **503 on `/owa/`** — OWA app pool crashing; check event logs for W3WP crash details
- **High auth latency in HttpProxy** — Domain controller connectivity or Kerberos token issues
- **Uneven backend distribution** — Load balancer misconfiguration or preferred server affinity

## Parallel Execution

The script uses PowerShell background jobs for parallel log collection:
- One job per IIS site per server
- One job per HttpProxy type per server
- All jobs run simultaneously across all servers
- Analysis runs after all jobs complete

For 2 servers with HttpProxy enabled, this starts ~16 parallel jobs (2 IIS sites × 2 servers + 6 proxy types × 2 servers).

## Related Tools

- **[Get-ExchangePerformanceDiagnostics.ps1](Get-ExchangePerformanceDiagnostics.ps1)** — Performance counters, DAG health, and Exchange metrics
- **[Export-Mailflow.ps1](Export-Mailflow.ps1)** — SMTP receive log analysis
