# Invoke-ExchangeOnlineDocumentation

Documents an Exchange Online tenant as an English "as-built" document: readable Markdown for Word, plus complete CSV exports with full detail.

The script is read-only — it never calls `Set-`/`New-`/`Remove-`/`Enable-`/`Disable-` cmdlets.

## Requirements

- Windows PowerShell 5.1
- `ExchangeOnlineManagement` module (v3) — must be installed beforehand (`Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser`)
- Exchange Online permissions: an **Exchange Administrator** or **Global Reader** role typically covers all cmdlets. For `-IncludePurview`, the account also needs Security & Compliance (Purview) access, e.g. **Compliance Administrator** or **Global Reader** + Purview role group.
- App-only authentication (`-AppId`/`-CertificateThumbprint`/`-Organization`) is supported for unattended runs; the app registration needs the `Exchange.ManageAsApp` permission and the appropriate Entra role.
- Some sections need additional roles or may show "Not available"/"None found" without them: address lists / address book policies (Address Lists role), Outlook add-ins (`Get-App`), room calendar processing, and priority accounts. Failures are listed in the appendix collection log.

## Usage

```powershell
.\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso"
.\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso" -IncludePurview
.\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso" -AppId $id -CertificateThumbprint $thumb -Organization contoso.onmicrosoft.com
```

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-CustomerName` | (mandatory) | Customer/tenant name, used in output file names |
| `-DocumentDate` | today | Date shown in the document information section |
| `-Author` | | Person producing the documentation |
| `-OutputPath` | `.\Reports` | Folder for the Markdown and CSV output |
| `-UserPrincipalName` | | UPN for the interactive connection |
| `-AppId` / `-CertificateThumbprint` / `-Organization` | | App-only certificate authentication (all three required together) |
| `-IncludePurview` | off | Also collect Security & Compliance (Purview) data |
| `-MaxListItems` | 10 | Cap for multi-valued properties rendered in-line |
| `-MaxTableRows` | 25 | Cap for overview tables; rest goes to CSV |
| `-CsvDelimiter` | culture | CSV delimiter override (default: `Export-Csv -UseCulture`) |
| `-DnsServer` | public DNS (1.1.1.1, then 8.8.8.8), falling back to the local resolver | DNS server for `Resolve-DnsName` lookups |
| `-ExoLogPath` | | Folder for ExchangeOnlineManagement client logs (`-EnableErrorReporting -LogLevel All`) |

## Output

```
<OutputPath>\EXO-Documentation_<Customer>_<yyyyMMdd>.md
<OutputPath>\EXO-Documentation_<Customer>_<yyyyMMdd>_csv\
```

The Markdown uses `#` = section, `##` = subsection, `###` = object card (`Setting | Value`, max 4 columns in any table). Each `#` section starts with a short description. Values are humanised: Yes/No, `yyyy-MM-dd` dates, "Not configured", lists capped with "… (+N more – see CSV `<file>`)".

Every dataset is exported to CSV in full — all rows, all collected properties. Datasets with zero rows get no file and appear as `none` in the appendix index.

## Document sections

1. Document information
2. Tenant overview
3. Organization settings
4. Domains and email authentication
5. Mail flow
6. Recipients
7. Groups
8. Client access and mailbox policies
9. Sharing and federation
10. Hybrid and migration
11. Email security (EOP / Defender for Office 365)
12. Permissions (RBAC)
13. Compliance and retention
14. Appendix (CSV index, collection log)

### What is documented

- **Organization settings** — general, transport and auditing flags, external sender tagging.
- **Domains** — accepted domains plus per-domain MX/SPF/DKIM (incl. DKIM selector CNAME check in DNS)/DMARC/MTA-STS/TLS-RPT.
- **Mail flow** — connectors, transport rules, remote domains, journal rules, HVE accounts.
- **Recipients** — counts and feature summary, mailbox plans, address lists and address book policies, public folders, room calendar processing.
- **Groups** — distribution/security/dynamic/M365 groups, M365 summary, groups without owners, groups accepting external senders, room lists.
- **Client access** — mailbox policy assignment spreads, protocol enablement, OWA/mobile/authentication policies, organization Outlook add-ins.
- **Sharing and federation** — organization relationships, sharing policies, availability address spaces.
- **Hybrid and migration** — on-premises organization, intra-org connectors, migration endpoints and batches.
- **Email security** — preset policies, all EOP/Defender threat policies with status/scope, quarantine policies incl. global notification settings, Tenant Allow/Block List, advanced delivery, user-reported settings, priority accounts, Teams protection.
- **RBAC** — role groups with roles and members, role assignment policies, custom roles, direct assignments with scopes, management scopes, RBAC for Applications.
- **Compliance** — MRM retention policies/tags, IRM/OME; with `-IncludePurview`: alert policies, Purview retention, DLP, sensitivity labels/label policies, retention labels.
- **Appendix** — CSV file index and the collection log (every failure is listed).

## Tests

```powershell
powershell -File .\Tests\Test-ExchangeOnlineDocumentation.ps1
```

The test harness extracts the script's functions, mocks every `Get-*` cmdlet, runs all sections and asserts the output — no tenant connection needed. Requires Windows PowerShell 5.1. Output goes to `%TEMP%\exo-doc-test`.

## Converting to Word (Apento template)

With pandoc (3.x) and the Apento reference template:

```powershell
& "C:\Users\<you>\AppData\Local\Pandoc\pandoc.exe" -f gfm `
    --reference-doc "C:\Users\<you>\OneDrive - APENTO\Apento\Apento word skabeloner\Løsnings design skabelon.docx" `
    -o EXO-Documentation.docx EXO-Documentation_<Customer>_<yyyyMMdd>.md
```

Alternatively, in Word: open the Apento template, place the cursor where the body should start, and use **Indsæt → Tekst fra fil** (Insert → Text from File) to pull the converted docx content into the template, keeping the cover page, versioning tables and TOC from the template.

The Markdown intentionally has no body title — the title goes on the template cover.
