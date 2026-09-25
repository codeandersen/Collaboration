# Exchange Online Review

## Overview

`Invoke-ExchangeOnlineReview.ps1` generates a read-only Markdown review report of an Exchange Online tenant. It collects Exchange Online, Exchange Online Protection (EOP) and Defender for Office 365 configuration without changing anything, and can optionally collect Microsoft Purview data, Microsoft Graph data, per-mailbox statistics and a message trace summary.

The customer is supplied with the `-CustomerName` parameter, so nothing tenant-specific is hard-coded. The report is data-only: it documents the current configuration with no ratings or recommendations.

### Output

- **Report**: `<OutputPath>\EXO-Review_<CustomerName>_<yyyyMMdd>.md` (UTF-8)
- **CSV export** (with `-ExportCsv`): `<OutputPath>\EXO-Review_<CustomerName>_<yyyyMMdd>_csv\<Section>.csv`
- Default output folder is `.\Reports`, which is git-ignored in this repo

## Prerequisites

- **PowerShell 7** recommended (`Resolve-DnsName` is used for the domain checks; Windows PowerShell 5.1 also works)
- **ExchangeOnlineManagement** module, latest v3:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
  ```
  The script installs it automatically if it's missing.
- **Microsoft.Graph** modules (only for `-IncludeGraph`):
  ```powershell
  Install-Module -Name Microsoft.Graph -Scope CurrentUser
  ```

## Permissions

### Exchange-only run (default)

- **Global Reader** is the minimum recommended role
- View-Only Organization Management or Security Reader also work
- Defender for Office 365 sections degrade gracefully to "Not available" when those cmdlets aren't licensed/permitted

### Purview run (`-IncludePurview`)

One of:

- Compliance Administrator
- Security Reader + Compliance Data Administrator
- View-Only Organization Management plus the Purview roles: View-Only DLP Compliance Management, View-Only Retention Management and Sensitivity Label Reader

Without `-IncludePurview` the script never calls `Connect-IPPSSession`; the Purview sections appear in the report marked "Skipped". With `-IncludePurview`, a failed IPPS connection marks the Purview sections "Not available" and the run continues — missing Purview permissions never block the EXO review.

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-CustomerName` | Yes | | Customer/tenant name; used in the report header and file name |
| `-ReviewedBy` | No | | Name of the reviewer, shown in the report header |
| `-ReviewDate` | No | Today | Review date shown in the report header |
| `-OutputPath` | No | `.\Reports` | Folder for the report and CSV exports (created if missing) |
| `-UserPrincipalName` | No | | UPN for the interactive EXO/IPPS sign-in |
| `-AppId` | No | | Application ID for app-only certificate authentication (requires `-CertificateThumbprint` and `-Organization`; used for both EXO and IPPS) |
| `-CertificateThumbprint` | No | | Certificate thumbprint for app-only authentication |
| `-Organization` | No | | Organization for app-only authentication (e.g. `contoso.onmicrosoft.com`) |
| `-IncludePurview` | No | Off | Connect to Security & Compliance PowerShell and add Purview sections |
| `-IncludeMailboxStatistics` | No | Off | Per-mailbox sizes, last logon and inbox rules forwarding externally. Slow on large tenants |
| `-IncludeMessageTrace` | No | Off | 7-day inbound/outbound message summary via `Get-MessageTraceV2` |
| `-IncludeGraph` | No | Off | Secure Score + Exchange-related controls and Exchange Administrator role holders via Microsoft Graph |
| `-ExportCsv` | No | Off | Export full detail tables as CSV files next to the report |
| `-MaxRows` | No | 50 | Maximum rows shown per Markdown table; longer tables are truncated with a "see CSV" note |
| `-DnsServer` | No | | Optional DNS server for `Resolve-DnsName` lookups |

## Examples

### EXO-only review (Purview sections show "Skipped")

```powershell
.\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso"
```

### Review including Purview data and CSV exports

```powershell
.\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" -IncludePurview -ExportCsv
```

### App-only certificate authentication

```powershell
.\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" `
    -AppId "00000000-0000-0000-0000-000000000000" `
    -CertificateThumbprint "A1B2C3..." `
    -Organization "contoso.onmicrosoft.com"
```

### Full deep scan

```powershell
.\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" `
    -IncludePurview -IncludeGraph -IncludeMailboxStatistics -IncludeMessageTrace -ExportCsv
```

## Report sections and cmdlets used

All sections run in the Exchange Online session unless marked otherwise. Every section is wrapped in its own error handling: a missing cmdlet, missing licence or access-denied error renders "Not available: \<reason\>" in the report and is recorded in the appendix — the run never aborts.

| Section | Cmdlets |
|---|---|
| 1. Organization configuration | `Get-OrganizationConfig`, `Get-ExternalInOutlook`, `Get-TransportConfig`, `Get-AdminAuditLogConfig` |
| 2. Recipients | `Get-EXOMailbox` (`-InactiveMailboxOnly`, `-SoftDeletedMailbox`), `Get-MailUser`, `Get-MailContact` |
| 3. Mailbox hygiene | `Get-EXOMailbox`, `Get-MailboxAuditBypassAssociation`, `Get-User`; opt-in: `Get-EXOMailboxStatistics`, `Get-InboxRule` |
| 4. Groups | `Get-DistributionGroup`, `Get-DynamicDistributionGroup`, `Get-UnifiedGroup` |
| 5. Domains and email authentication | `Get-AcceptedDomain`, `Get-DkimSigningConfig`, `Get-ArcConfig`, `Resolve-DnsName` (MX, SPF, DMARC, MTA-STS, TLS-RPT, BIMI) |
| 6. Mail flow | `Get-InboundConnector` (Enhanced Filtering), `Get-OutboundConnector`, `Get-TransportRule`, `Get-RemoteDomain`, `Get-JournalRule`, `Get-MailUser -HVEAccount`; opt-in: `Get-MessageTraceV2` (summarised by direction and status) |
| 7. Hybrid and migration | `Get-OnPremisesOrganization`, `Get-IntraOrganizationConnector`, `Get-MigrationEndpoint`, `Get-MigrationBatch` |
| 8. Sharing | `Get-OrganizationRelationship`, `Get-SharingPolicy`, `Get-AvailabilityAddressSpace` |
| 9. Client access | `Get-OwaMailboxPolicy`, `Get-ActiveSyncOrganizationSettings`, `Get-ActiveSyncDeviceAccessRule`, `Get-MobileDeviceMailboxPolicy`, `Get-EXOCASMailbox`, `Get-AuthenticationPolicy` |
| 10. Permissions and RBAC | `Get-RoleGroup`, `Get-RoleGroupMember`, `Get-RoleAssignmentPolicy`, `Get-ManagementRole`, `Get-ManagementRoleAssignment`, `Get-ManagementScope`, `Get-ServicePrincipal`; opt-in Graph: `Get-MgDirectoryRole(Member)` |
| 11. Threat protection (EOP / Defender) | `Get-AntiPhishPolicy`/`Rule`, `Get-HostedContentFilterPolicy`/`Rule`, `Get-HostedConnectionFilterPolicy`, `Get-HostedOutboundSpamFilterPolicy`, `Get-MalwareFilterPolicy`/`Rule`, `Get-SafeAttachmentPolicy`/`Rule`, `Get-AtpPolicyForO365`, `Get-SafeLinksPolicy`/`Rule`, `Get-EOPProtectionPolicyRule`, `Get-ATPProtectionPolicyRule`, `Get-ATPBuiltInProtectionRule`, `Get-TenantAllowBlockListItems`, `Get-TenantAllowBlockListSpoofItems`, `Get-SecOpsOverridePolicy`, `Get-ExoSecOpsOverrideRule`, `Get-PhishSimOverridePolicy`, `Get-ExoPhishSimOverrideRule`, `Get-QuarantinePolicy`, `Get-EmailTenantSettings`, `Get-ReportSubmissionPolicy`, `Get-TeamsProtectionPolicy` |
| 12. Alert policies (Purview) | `Get-ProtectionAlert` |
| 13. Compliance and retention | EXO: `Get-RetentionPolicy`, `Get-RetentionPolicyTag`, `Get-IRMConfiguration`, `Get-OMEConfiguration`. Purview: `Get-RetentionCompliancePolicy`/`Rule`, `Get-DlpCompliancePolicy`/`Rule`, `Get-Label`, `Get-LabelPolicy`, `Get-ComplianceTag` |
| 14. Appendix | Collection log (skipped/failed sections with reasons), parameters used, run duration, cmdlet reference |

The header of the report records which data sources were collected (Purview: Yes / No / Failed, Graph, mailbox statistics, message trace) and includes a table of contents and a summary counts table.

## Troubleshooting

**Error: "Failed to connect to Exchange Online"**
- Check network/proxy access to `outlook.office365.com`
- Make sure the account has Exchange Online access and modern auth works (MFA prompt expected)
- For app-only auth, verify the app registration has `Exchange.ManageAsApp` (Office 365 Exchange Online API) and the certificate is installed in the user/machine store

**Warning: "Could not connect to Security & Compliance PowerShell"**
- Expected when the account lacks Purview permissions — the EXO review still completes and the Purview sections are marked "Not available". See the Permissions section above for the roles needed.

**Sections show "Not available: cmdlet ... not available"**
- Defender for Office 365 sections require an E5/Defender P1 or P2 licence. The rest of the report is unaffected.
- Update the ExchangeOnlineManagement module: `Update-Module ExchangeOnlineManagement`

**DNS checks return empty values**
- `Resolve-DnsName` requires the records to be publicly resolvable. Use `-DnsServer` to specify a resolver (e.g. `8.8.8.8`) if the machine's default resolver blocks lookups.
- DNS checks are skipped for `*.onmicrosoft.com` domains (only DKIM/DMARC are checked for the MOERA domain).

**"Microsoft.Graph modules not found"**
- Install with `Install-Module Microsoft.Graph -Scope CurrentUser`, or drop `-IncludeGraph`.

**Slow runs**
- `-IncludeMailboxStatistics` runs `Get-EXOMailboxStatistics` (and `Get-InboxRule`) per mailbox — expect several minutes on large tenants.
- `-IncludeMessageTrace` is throttled server-side; the report notes the sample size.

## Data sensitivity

The report contains tenant configuration details, user names, email addresses and group memberships. Treat it as confidential. The default `Reports/` output folder is excluded from git via `.gitignore`; review the report before sharing it outside the customer's team.

## Read-only guarantee

The script only calls `Get-*` cmdlets, `Resolve-DnsName`, `Connect-ExchangeOnline`, `Connect-IPPSSession`, `Connect-MgGraph` and the matching `Disconnect-*` cmdlets. It never calls `Set-`, `New-`, `Remove-`, `Enable-` or `Disable-` Exchange cmdlets, and makes no changes to the tenant.

## Version History

- **v1.0** (2026-09-24): Initial release
  - Markdown report covering EXO, EOP and Defender for Office 365 configuration
  - Opt-in Purview, Graph, mailbox statistics and message trace sections
  - Optional CSV export

## License

This script is provided as-is for Exchange Online tenant reviews. Test against a non-production tenant before using it in production.
