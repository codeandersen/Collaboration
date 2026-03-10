# Exchange Online Edge Blocking Readiness Assessment

## Overview

This toolkit helps you assess whether your Exchange hybrid environment is ready to enable **edge blocking** in Exchange Online. Edge blocking means all inbound mail flows directly to Exchange Online, bypassing on-premises Exchange servers, providing better protection and improved mail flow.

## What is Edge Blocking?

**Edge blocking** (also called centralized mail transport) routes all inbound internet mail directly to Exchange Online instead of your on-premises Exchange servers.

### Benefits
- ✅ **Better Protection**: Microsoft's advanced threat protection filters all inbound mail
- ✅ **Improved Mail Flow**: Direct delivery to cloud mailboxes
- ✅ **Reduced Complexity**: Simplified mail flow architecture
- ✅ **Lower Costs**: Reduced on-premises infrastructure requirements

### Requirements
- All mail-enabled recipients must exist in Exchange Online (as mailboxes, mail users, or groups)
- MX records point to Exchange Online (*.mail.protection.outlook.com)
- Mail-enabled public folders must be synchronized as MailUser objects

## 🚨 Important: Public Folders

**Good News**: You **CAN** enable edge blocking even with on-premises public folders!

**Requirements**:
- Mail-enabled public folders must be synchronized to Exchange Online as **MailUser** objects
- Use Microsoft's `Sync-MailPublicFolders.ps1` script: https://aka.ms/SyncMailPublicFolders
- Public folder content stays on-premises (no migration needed)
- Only mail routing changes

## Scripts Included

### 1. Get-Exchange2016MailObjects.ps1
Inventories all mail-enabled objects in Exchange 2016 on-premises.

**Run on**: Exchange 2016 server or workstation with Exchange Management Shell

**Output**: `C:\Temp\Exchange2016-MailEnabledObjects.csv`

**Collects**:
- User mailboxes
- Remote mailboxes (already in cloud)
- Distribution groups
- Mail-enabled security groups
- Dynamic distribution groups
- Mail contacts
- Mail users
- Mail-enabled public folders

### 2. Get-ExchangeOnlineMailObjects.ps1
Inventories all mail-enabled objects in Exchange Online.

**Run on**: Workstation with Exchange Online PowerShell module

**Output**: `C:\Temp\ExchangeOnline-MailEnabledObjects.csv`

**Collects**:
- User mailboxes
- Distribution groups
- Mail-enabled security groups
- Dynamic distribution groups
- Mail contacts
- Mail users
- Microsoft 365 groups

### 3. Compare-EdgeBlockingReadiness.ps1
Compares both environments and generates a comprehensive HTML readiness report.

**Run on**: Any workstation with both CSV files

**Output**: `C:\Temp\EdgeBlocking-ReadinessReport.html`

**Analyzes**:
- Objects missing in Exchange Online
- On-premises mailboxes not yet migrated
- Public folder synchronization status
- Overall readiness for edge blocking

## Prerequisites

### For Exchange 2016 Script
- Exchange Management Shell access
- Permissions to read all recipient objects
- Network access to all Exchange 2016 servers

### For Exchange Online Script
- Exchange Online PowerShell module v3.0 or later
  ```powershell
  Install-Module -Name ExchangeOnlineManagement
  ```
- Global Admin or Exchange Administrator role
- Modern authentication enabled

### For Comparison Script
- Both CSV files from previous scripts
- Web browser to view HTML report

## Quick Start

### Step 1: Run on Exchange 2016 Server
```powershell
# Navigate to the script directory
cd "C:\Github\Collaboration\Exchange Hybrid"

# Run the on-premises inventory
.\Get-Exchange2016MailObjects.ps1
```

### Step 2: Run on Workstation with Exchange Online Access
```powershell
# Navigate to the script directory
cd "C:\Github\Collaboration\Exchange Hybrid"

# Run the cloud inventory
.\Get-ExchangeOnlineMailObjects.ps1
```

### Step 3: Generate Comparison Report
```powershell
# Navigate to the script directory
cd "C:\Github\Collaboration\Exchange Hybrid"

# Run the comparison and generate report
.\Compare-EdgeBlockingReadiness.ps1
```

The HTML report will automatically open in your default browser.

## Custom Export Paths

You can specify custom paths for exports:

```powershell
# Custom on-premises export
.\Get-Exchange2016MailObjects.ps1 -ExportPath "C:\Reports\OnPrem-Recipients.csv"

# Custom cloud export
.\Get-ExchangeOnlineMailObjects.ps1 -ExportPath "C:\Reports\Cloud-Recipients.csv"

# Custom report with custom source files
.\Compare-EdgeBlockingReadiness.ps1 `
    -OnPremCsvPath "C:\Reports\OnPrem-Recipients.csv" `
    -CloudCsvPath "C:\Reports\Cloud-Recipients.csv" `
    -ReportPath "C:\Reports\EdgeBlocking-Assessment.html"
```

## Understanding the Report

### Status Indicators

- **✅ YES (Green)**: Ready to enable edge blocking
- **❌ NO (Red)**: Blockers must be resolved first

### Common Blockers

1. **On-Premises Mailboxes Not Migrated**
   - **Impact**: Critical blocker
   - **Resolution**: Migrate mailboxes to Exchange Online
   - **Methods**: Hybrid migration, cutover, or staged migration

2. **Mail-Enabled Objects Missing in Cloud**
   - **Impact**: Critical blocker
   - **Resolution**: Ensure Azure AD Connect is syncing all objects
   - **Check**: Sync scope, OU filters, and sync errors

3. **Public Folders Not Synchronized**
   - **Impact**: Critical blocker (if you have mail-enabled public folders)
   - **Resolution**: Run `Sync-MailPublicFolders.ps1`
   - **Note**: Content stays on-premises, only mail routing changes

### Warnings vs Blockers

- **Blockers**: Must be fixed before enabling edge blocking
- **Warnings**: Review and address, but may not prevent edge blocking

## Public Folder Synchronization

If you have mail-enabled public folders, follow these steps:

### 1. Download the Sync Script
```powershell
# Download from Microsoft
Invoke-WebRequest -Uri "https://aka.ms/SyncMailPublicFolders" -OutFile ".\Sync-MailPublicFolders.ps1"
```

### 2. Run the Sync Script
```powershell
# Connect to both environments
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Connect-ExchangeOnline

# Run the sync
.\Sync-MailPublicFolders.ps1 -Credential (Get-Credential) -CsvSummaryFile ".\PFSync-Summary.csv"
```

### 3. Verify Synchronization
```powershell
# Check that MailUser objects were created
Get-MailUser | Where-Object {$_.RecipientTypeDetails -eq 'MailUser' -and $_.EmailAddresses -like '*publicfolder*'}
```

### 4. Test Mail Flow
Send test emails to mail-enabled public folders and verify delivery.

## Enabling Edge Blocking

Once the report shows **✅ YES**, follow these steps:

### 1. Backup Current Configuration
```powershell
# Document current MX records
nslookup -type=MX yourdomain.com

# Export current connectors
Get-InboundConnector | Export-Clixml ".\Backup-InboundConnectors.xml"
Get-OutboundConnector | Export-Clixml ".\Backup-OutboundConnectors.xml"
```

### 2. Update MX Records
Point your MX records to Exchange Online:
- **Priority**: 0
- **Host**: yourdomain-com.mail.protection.outlook.com
- **TTL**: 3600 (or your standard)

### 3. Configure Inbound Connector (Optional)
If you need specific routing for certain scenarios:
```powershell
New-InboundConnector -Name "From On-Premises" `
    -ConnectorType OnPremises `
    -SenderDomains * `
    -RequireTls $true `
    -RestrictDomainsToIPAddresses $true `
    -SenderIPAddresses "your-onprem-ip"
```

### 4. Test Mail Flow
```powershell
# Test inbound mail to various recipient types
# - User mailboxes
# - Distribution groups
# - Shared mailboxes
# - Public folders (if applicable)
# - Mail contacts
```

### 5. Monitor
```powershell
# Check message trace for delivery issues
Get-MessageTrace -StartDate (Get-Date).AddHours(-24) -EndDate (Get-Date) | 
    Where-Object {$_.Status -ne 'Delivered'} | 
    Format-Table Received, SenderAddress, RecipientAddress, Status
```

### 6. Watch for NDRs
Monitor for 48 hours and address any non-delivery reports.

## Troubleshooting

### Script Errors

**Error**: "Exchange Management Shell not loaded"
- **Solution**: Run from Exchange Management Shell or load the snap-in:
  ```powershell
  Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
  ```

**Error**: "Get-EXOMailbox not found"
- **Solution**: Install Exchange Online PowerShell module:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Force
  Connect-ExchangeOnline
  ```

### Common Issues

**Issue**: Objects showing as missing but they exist
- **Check**: Case sensitivity in email addresses
- **Check**: Proxy addresses vs primary SMTP address
- **Solution**: Review the CSV files manually

**Issue**: Public folders not syncing
- **Check**: Mail-enabled status in on-premises
- **Check**: Permissions to create MailUser objects in Exchange Online
- **Solution**: Run sync script with verbose logging

**Issue**: Azure AD Connect not syncing objects
- **Check**: Sync scope and OU filters
- **Check**: Sync errors in Azure AD Connect Health
- **Solution**: Review and adjust sync configuration

## Best Practices

1. **Run During Maintenance Window**: Schedule assessment during low-usage periods
2. **Test Mail Flow Thoroughly**: Test all recipient types before going live
3. **Keep Backups**: Export all connector and MX record configurations
4. **Monitor Closely**: Watch message trace for 48 hours after enabling
5. **Document Changes**: Keep detailed records of all configuration changes
6. **Communicate**: Inform users about the change and potential impact

## Rollback Plan

If you need to revert edge blocking:

1. **Restore MX Records**: Point back to on-premises servers
2. **Wait for DNS Propagation**: Allow TTL time to expire
3. **Restore Connectors**: Import backed-up connector configurations
4. **Test Mail Flow**: Verify mail delivery to all recipient types
5. **Document Issues**: Record what went wrong for future attempts

## Additional Resources

- [Microsoft Learn: Exchange Hybrid Deployments](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/deploy-hybrid)
- [Manage mail flow using a third-party cloud service](https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/manage-mail-flow-using-third-party-cloud)
- [Synchronize mail-enabled public folders](https://learn.microsoft.com/en-us/exchange/collaboration-exo/public-folders/sync-mail-public-folders)
- [Mailbox migrations in Exchange Online](https://learn.microsoft.com/en-us/exchange/mailbox-migration/mailbox-migration)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)

## Support

For issues with:
- **Scripts**: Review error messages and check prerequisites
- **Migrations**: Consult Microsoft migration guides
- **Public Folders**: Use Microsoft's Sync-MailPublicFolders.ps1 script
- **Azure AD Connect**: Check Azure AD Connect Health portal
- **Mail Flow**: Use Exchange Online message trace and mail flow troubleshooter

## Version History

- **v1.0** (2026-03-10): Initial release
  - Exchange 2016 inventory script
  - Exchange Online inventory script
  - Comparison and reporting script
  - Public folder support
  - Comprehensive HTML reporting

## License

These scripts are provided as-is for Exchange hybrid environment assessment. Test thoroughly in a non-production environment before using in production.

---

**Created**: March 10, 2026  
**Author**: Exchange Hybrid Assessment Toolkit  
**Purpose**: Assess readiness for Exchange Online edge blocking
