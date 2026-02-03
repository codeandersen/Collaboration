# Exchange Online Retention Policy and Archive Enablement - Documentation

## Overview

This document provides comprehensive documentation for the `set_retention_and_enable_archive.ps1` script, which automates the assignment of retention policies and enablement of mailbox archiving in Exchange Online for large-scale environments.

**Last Updated**: January 21, 2026

## Azure Automation Configuration

### Subscription Details
- **Subscription**: `sub-starkgrp-compliance-prod`
- **Automation Account**: `aa-compliance-mail-archive-retention`
- **Runbook Name**: `set_retention_and_enable_archive`

### Managed Identity Configuration

The automation account uses a **System Managed Identity** (Service Principal) with the following configuration:

- **Object ID**: `163e4418-967c-4841-9d32-4a5ca5f89e51`
- **Azure AD Role**: Exchange Administrator
- **Graph API Permission**: `Exchange.ManageAsApp` (Application permission)

### Required Permissions

The managed identity requires the following permissions to function correctly:

1. **Exchange Administrator Role** (Azure AD)
   - Allows full management of Exchange Online resources
   - Required for mailbox operations and retention policy assignment

2. **Microsoft Graph API - Exchange.ManageAsApp**
   - Application-level permission for Exchange operations via Graph API
   - Required for querying user licenses and mailbox information

## Script Purpose

The script performs two primary functions across all mailboxes in the Exchange Online tenant:

1. **Retention Policy Assignment**: Applies the specified retention policy to eligible mailboxes
2. **Archive Enablement**: Enables in-place archiving for eligible mailboxes

## Key Features

### Performance Optimization (Updated)

**Major Performance Improvements**:
- **Bulk License Fetching**: All user licenses are fetched upfront in a single bulk operation, eliminating thousands of individual Graph API calls
- **Pre-filtering Mailboxes**: Only processes mailboxes that need attention (missing retention policy or inactive archives)
- **Improved Archive Detection**: Uses `ArchiveGuid` property for more reliable archive status checking
- **Memory Management**: Implements garbage collection every 500 mailboxes to prevent memory exhaustion
- **Progress Reporting**: Provides detailed progress updates every 500 mailboxes with processing rate metrics

### Error Handling
- **Retry Logic**: Automatically retries transient errors up to 3 times with 2-second delays
- **Error Classification**: Distinguishes between transient and permanent errors
- **Detailed Logging**: Comprehensive error messages with user identification
- **Null Safety**: Enhanced null checking for error messages and mailbox properties

### License Validation
- **Bulk License Lookup**: Pre-loads all user licenses into memory for instant lookups
- **Multiple SKU Support**: Validates against approved license types
- **Shared Mailbox Logic**: Special handling for unlicensed and licensed shared mailboxes

## Script Parameters

| Parameter | Type | Default Value | Description |
|-----------|------|---------------|-------------|
| `ExemptSecurityGroup` | String | `Exempt_OnlineArchive@starkworkspace.onmicrosoft.com` | Mail-enabled security group for mailboxes exempt from archiving |
| `RetentionPolicyName` | String | `STARK Group Default` | Name of the retention policy to apply |
| `WhatIf` | Switch | `$true` | Runs in preview mode without making changes when enabled |

## Exempt Group Behavior

Mailboxes that are members of the exempt security group (`Exempt_OnlineArchive@starkworkspace.onmicrosoft.com`) receive special handling:

- ✅ **Retention Policy**: Applied normally
- ❌ **Archive Enablement**: Skipped (not enabled)

This allows organizations to apply retention policies to all users while selectively controlling which mailboxes have archiving enabled.

## Valid License SKUs

The script validates that users have one of the following license SKUs before processing:

| SKU Part Number | License Name |
|-----------------|--------------|
| `SPE_E5` | Microsoft 365 E5 |
| `SPE_E3` | Microsoft 365 E3 |
| `SPE_F5` | Microsoft 365 F5 |
| `SPE_F5_SECCOMP` | Microsoft 365 F5 Security & Compliance |
| `ENTERPRISEPREMIUM` | Office 365 E5 |
| `ENTERPRISEPACK` | Office 365 E3 |
| `EXCHANGEENTERPRISE` | Exchange Online Plan 2 |
| `EXCHANGEARCHIVE_ADDON` | Exchange Online Archiving |

### Shared Mailbox License Handling

Shared mailboxes receive special license validation logic:

1. **Unlicensed Shared Mailboxes**: Automatically approved (valid)
2. **Exchange Online Plan 2**: Automatically approved
3. **Exchange Online Plan 1 + Archiving Add-on**: Automatically approved
4. **Other Valid SKUs**: Approved if licensed with any SKU from the valid list
5. **Invalid Licenses**: Skipped with detailed logging

## Processing Workflow

### 1. Initialization Phase
```
1. Connect to Azure using Managed Identity
2. Obtain Graph API access token
3. Connect to Exchange Online using Managed Identity
4. Retrieve organization SKU information
5. Bulk-fetch ALL user licenses (new optimization)
6. Build exempt mailbox lookup table
7. Pre-filter mailboxes needing attention (new optimization)
```

### 2. Mailbox Processing Phase
```
For each mailbox that needs processing:
  1. Check if exempt from archiving
  2. Lookup user licenses from pre-loaded cache (no API call)
  3. Validate license eligibility
  4. Apply retention policy (with retry logic)
  5. Enable archive (if not exempt and not already enabled)
  6. Log results and update counters
```

### 3. Completion Phase
```
1. Display final summary statistics
2. Disconnect from Exchange Online
3. Disconnect from Azure Account
```

## Performance Improvements

### Bulk License Fetching (New)

**Lines 72-91**: The script now fetches all user licenses in bulk at startup:
- Uses Graph API pagination with `$top=999` for optimal performance
- Stores licenses in a hashtable for O(1) lookup speed
- Progress reporting every 5,000 users
- Eliminates individual API calls per mailbox (massive performance gain)

### Pre-filtering Mailboxes (New)

**Lines 125-129**: Only processes mailboxes that need attention:
```powershell
Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum,Retention,Archive | Where-Object {
    $_.RetentionPolicy -ne $RetentionPolicyName -or 
    $_.ArchiveGuid -eq '00000000-0000-0000-0000-000000000000'
}
```

This filters out mailboxes that already have the correct retention policy AND active archives, significantly reducing processing time.

### Improved Archive Detection (New)

**Line 263**: Now uses `ArchiveGuid` instead of `ArchiveStatus`:
```powershell
if ($mbx.ArchiveGuid -ne '00000000-0000-0000-0000-000000000000')
```

This provides more reliable archive status detection.

## Output and Reporting

### Progress Reports (Every 500 Mailboxes)
- Total mailboxes processed
- Retention policies assigned
- Retention policies already set
- Archives enabled
- Archives already active
- Exempt mailboxes skipped
- Unlicensed mailboxes skipped
- Error counts
- Processing rate (mailboxes/second)

### Final Summary
The script provides a comprehensive summary including:
- Total mailboxes processed (only those needing attention)
- Exempt mailboxes (retention applied, archive skipped)
- Unlicensed mailboxes (skipped entirely)
- Retention policy assignment results
- Archive enablement results
- Error statistics

## Error Handling

### Transient Errors
Errors matching patterns like "server side error" or "try again after some time" are automatically retried:
- Maximum 3 retry attempts
- 2-second delay between retries
- Logged separately in final statistics

### Permanent Errors
All other errors are logged immediately and counted separately:
- Retention policy assignment errors
- Archive enablement errors
- License query errors (now rare due to bulk fetching)

### Enhanced Null Safety (New)

**Lines 290-293**: Improved error handling with null checks:
```powershell
$errorMsg = if ($_.Exception.Message) { $_.Exception.Message } else { "Unknown error" }
$mbxDisplay = if ($mailbox.DisplayName) { $mailbox.DisplayName } else { "Unknown" }
$mbxEmail = if ($mailbox.PrimarySmtpAddress) { $mailbox.PrimarySmtpAddress } else { "Unknown" }
```

## WhatIf Mode

By default, the script runs in **WhatIf mode** (`$WhatIf = $true`):
- No actual changes are made to mailboxes
- All operations are simulated and logged
- Allows validation before production execution
- Set to `$false` to execute actual changes

## Prerequisites

### Required PowerShell Modules
- `ExchangeOnlineManagement`
- `Az.Accounts`

### Network Requirements
- Connectivity to Microsoft Graph API (`https://graph.microsoft.com`)
- Connectivity to Exchange Online endpoints

### Permissions Summary
The managed identity must have:
1. Exchange Administrator role in Azure AD
2. Exchange.ManageAsApp permission in Microsoft Graph

## Monitoring and Troubleshooting

### Common Issues

**Issue**: Slow performance with large tenant
- **Resolution**: Script now uses bulk license fetching and pre-filtering (resolved in current version)

**Issue**: Memory exhaustion in large environments
- **Cause**: Processing too many mailboxes without cleanup
- **Resolution**: Garbage collection runs every 500 mailboxes (line 302)

**Issue**: Transient connection errors
- **Cause**: Temporary Exchange Online service issues
- **Resolution**: Automatic retry logic with 3 attempts and 2-second delays

**Issue**: Archive not detected correctly
- **Resolution**: Now uses `ArchiveGuid` for reliable detection (line 263)

### Logging
All output is written to the Azure Automation job output stream:
- Standard output for normal operations
- Warning output for non-critical issues
- Error output for failures

## Maintenance

### Updating the Exempt Group
To add or remove users from the exempt group, use the companion script:
- **Script**: `Add-UsersToExemptGroup.ps1`
- **Purpose**: Bulk add users to the exempt security group from CSV
- **Location**: Same directory as main script

### Updating Valid License SKUs
To modify the list of valid licenses, update line 153:
```powershell
$validSkus = @('SPE_E5', 'SPE_E3', 'SPE_F5', ...)
```

### Changing Retention Policy
To apply a different retention policy, modify the parameter:
```powershell
-RetentionPolicyName "New Policy Name"
```

## Security Considerations

1. **Managed Identity**: Uses Azure Managed Identity for authentication (no stored credentials)
2. **Least Privilege**: Service principal has only required permissions
3. **Audit Trail**: All operations are logged in Azure Automation job history
4. **WhatIf Default**: Prevents accidental changes with default WhatIf mode

## Performance Metrics

### Expected Performance
With the new optimizations:
- **Bulk License Fetch**: ~5,000-10,000 users per minute
- **Mailbox Processing**: Varies based on operations needed
- **Memory Usage**: Stable with periodic garbage collection
- **API Calls**: Dramatically reduced (no per-mailbox Graph API calls)

### Scalability
The script is optimized for environments with:
- 25,000+ mailboxes
- Large numbers of licenses
- Mixed mailbox types (user, shared)

## Version History

### Version 2.0 (January 2026)
- **Major Performance Improvements**:
  - Bulk license fetching eliminates thousands of individual API calls
  - Pre-filtering only processes mailboxes needing attention
  - Improved archive detection using `ArchiveGuid`
- **Enhanced Error Handling**:
  - Better null safety for error messages
  - More robust error reporting
- **Code Quality**:
  - Removed obsolete special character escaping (no longer needed with bulk fetch)
  - Cleaner code structure

### Version 1.0 (January 2026)
- Initial version with streaming processing
- Special character handling for UPNs with apostrophes
- Exempt group receives retention policy but not archiving
- Support for shared mailbox license validation
- Enhanced error handling and retry logic

## Support and Contact

For issues or questions regarding this automation:
1. Review Azure Automation job output logs
2. Check managed identity permissions
3. Verify Exchange Online connectivity
4. Review Graph API permissions

## Related Scripts

- **Add-UsersToExemptGroup.ps1**: Bulk add users to exempt security group
- **verify-archive-retention.ps1**: Verify archive and retention policy status
- **Set-MailboxArchive.ps1**: Individual mailbox archive management

## Technical Notes

### Why Bulk License Fetching?
The previous version made individual Graph API calls for each mailbox to check licenses. With 25,000+ mailboxes, this resulted in:
- 25,000+ API calls
- Significant processing time
- Potential throttling issues

The new bulk approach:
- Makes ~25-50 API calls total (depending on tenant size)
- Loads all licenses into memory once
- Provides instant lookups during processing
- Dramatically improves performance

### Why Pre-filtering?
In a mature environment, most mailboxes already have the correct configuration. Pre-filtering:
- Skips fully configured mailboxes at the query level
- Reduces processing time
- Focuses effort on mailboxes needing attention
- Still provides accurate reporting

### Archive Detection Method
Using `ArchiveGuid` instead of `ArchiveStatus`:
- More reliable indicator of archive presence
- GUID of all zeros = no archive
- Non-zero GUID = archive exists
- Avoids status interpretation issues
