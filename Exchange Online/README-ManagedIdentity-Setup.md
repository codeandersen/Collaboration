# Shared Mailbox Permission Management - Azure Automation Setup Guide

## Overview

This guide walks you through setting up automated shared mailbox permission management using **Azure Automation with Managed Identity**. The script automatically grants/revokes FullAccess (with automapping) and SendAs permissions based on distribution group membership.

### How It Works

- **Group Naming Convention**: `SMBX-<mailboxname>`
- **Example**: Group `SMBX-sales` manages permissions for shared mailbox `sales@contoso.com`
- **Automatic Sync**: Users added to the group get permissions; users removed lose permissions
- **Automapping Enabled**: Shared mailboxes automatically appear in Outlook

---

## Prerequisites

- **Azure Subscription** with permissions to create Automation Accounts
- **Microsoft 365 tenant** with Exchange Online
- **Global Administrator** or **Privileged Role Administrator** access
- **Exchange Administrator** role assignment capability

---

## Step-by-Step Setup

### Step 1: Create Azure Automation Account

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Create a resource** → Search for **Automation**
3. Click **Create** on **Automation Account**
4. Configure:
   - **Subscription**: Select your subscription
   - **Resource Group**: Create new or use existing
   - **Name**: `AutomationAccount-SharedMailbox` (or your preferred name)
   - **Region**: Select closest region
   - **System assigned managed identity**: **Enable** ✅
5. Click **Review + Create** → **Create**
6. Wait for deployment to complete

---

### Step 2: Enable and Configure Managed Identity

1. Open your newly created Automation Account
2. Navigate to **Account Settings** → **Identity**
3. Under **System assigned** tab:
   - Status should be **On** (if not, turn it on)
   - Copy the **Object (principal) ID** - you'll need this

---

### Step 3: Grant Exchange Administrator Role to Managed Identity

#### Option A: Using Azure Portal

1. Navigate to [Azure Active Directory](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade)
2. Click **Roles and administrators**
3. Search for and click **Exchange Administrator**
4. Click **+ Add assignments**
5. Search for your Automation Account name (e.g., `AutomationAccount-SharedMailbox`)
6. Select it and click **Add**
7. Verify the assignment appears in the list

#### Option B: Using PowerShell

```powershell
# Connect to Microsoft Graph
Connect-MgGraph -Scopes "RoleManagement.ReadWrite.Directory"

# Get the Managed Identity Object ID (replace with your Object ID from Step 2)
$ManagedIdentityObjectId = "YOUR-OBJECT-ID-HERE"

# Get Exchange Administrator role
$ExchangeAdminRole = Get-MgDirectoryRole -Filter "DisplayName eq 'Exchange Administrator'"

# If role not activated, activate it first
if (-not $ExchangeAdminRole) {
    $RoleTemplate = Get-MgDirectoryRoleTemplate -Filter "DisplayName eq 'Exchange Administrator'"
    $ExchangeAdminRole = New-MgDirectoryRole -RoleTemplateId $RoleTemplate.Id
}

# Assign role to Managed Identity
New-MgDirectoryRoleMemberByRef -DirectoryRoleId $ExchangeAdminRole.Id -BodyParameter @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$ManagedIdentityObjectId"
}

Write-Host "Exchange Administrator role assigned successfully!" -ForegroundColor Green
```

---

### Step 4: Install ExchangeOnlineManagement Module

1. In your Automation Account, navigate to **Shared Resources** → **Modules**
2. Click **+ Add a module**
3. Select **Browse from gallery**
4. Search for `ExchangeOnlineManagement`
5. Select the module and click **Import**
6. Wait for import to complete (Status: **Available**)

**Note**: Ensure you have version **3.0.0 or higher**

---

### Step 5: Create the Runbook

1. In your Automation Account, navigate to **Process Automation** → **Runbooks**
2. Click **+ Create a runbook**
3. Configure:
   - **Name**: `Sync-SharedMailboxPermissions`
   - **Runbook type**: **PowerShell**
   - **Runtime version**: **7.2** (or latest)
   - **Description**: `Automatically syncs shared mailbox permissions based on SM- group membership`
4. Click **Create**

---

### Step 6: Add the Script Code

1. Open the newly created runbook
2. Click **Edit**
3. Copy the entire content from `Sync-SharedMailboxPermissions.ps1`
4. Paste into the editor
5. Click **Save**
6. Click **Publish**
7. Confirm by clicking **Yes**

---

### Step 7: Test the Runbook

Before scheduling, test manually:

1. Create a test shared mailbox (e.g., `testmailbox@contoso.com`)
2. Create a distribution group named `SM-testmailbox@contoso.com`
3. Add a test user to the group
4. In the runbook, click **Start**
5. Optionally add parameters:
   - **GROUPPREFIX**: `SM-` (default)
   - **AUTHMETHOD**: `ManagedIdentity` (default)
6. Click **OK**
7. Monitor the **Output** tab for results
8. Verify the test user received permissions on the shared mailbox

#### Verification Commands

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Check permissions
Get-MailboxPermission -Identity "testmailbox@contoso.com" | 
    Where-Object {$_.User -notlike "NT AUTHORITY*"} | 
    Format-Table User, AccessRights, AutoMapping

Get-RecipientPermission -Identity "testmailbox@contoso.com" | 
    Where-Object {$_.Trustee -notlike "NT AUTHORITY*"} | 
    Format-Table Trustee, AccessRights
```

---

### Step 8: Schedule the Runbook

1. In the runbook, click **Schedules** → **+ Add a schedule**
2. Click **Link a schedule to your runbook**
3. Click **+ Add a schedule**
4. Configure:
   - **Name**: `Daily-SharedMailbox-Sync`
   - **Description**: `Runs daily to sync shared mailbox permissions`
   - **Starts**: Tomorrow at 2:00 AM (or preferred time)
   - **Time zone**: Select your timezone
   - **Recurrence**: **Recurring**
   - **Recur every**: `1 Day`
   - **Expiration**: Never expire (or set as needed)
5. Click **Create**
6. Configure parameters:
   - **GROUPPREFIX**: `SM-`
   - **AUTHMETHOD**: `ManagedIdentity`
7. Click **OK**

**Recommended Schedule**: Run daily during off-peak hours (e.g., 2:00 AM)

---

## Usage Guide

### Creating Shared Mailbox Access Groups

1. **Create the shared mailbox** (if not exists):
   ```powershell
   New-Mailbox -Shared -Name "Sales Team" -Alias "sales" -PrimarySmtpAddress "sales@contoso.com"
   ```

2. **Create the distribution group** with `SMBX-` prefix:
   ```powershell
   New-DistributionGroup -Name "SMBX-sales@contoso.com" -Alias "SMBX-sales" -Type "Security" -ManagedBy "admin@contoso.com"
   ```

3. **Add users to the group**:
   ```powershell
   Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "john.doe@contoso.com"
   Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "jane.smith@contoso.com"
   ```

4. **Wait for next scheduled run** or **manually trigger** the runbook

5. **Verify permissions** were applied

### Naming Convention Examples

| Distribution Group | Shared Mailbox | Result |
|-------------------|----------------|--------|
| `SMBX-sales@contoso.com` | `sales@contoso.com` | ✅ Match |
| `SMBX-support@contoso.com` | `support@contoso.com` | ✅ Match |
| `SMBX-hr@contoso.com` | `hr@contoso.com` | ✅ Match |
| `SMBX-finance` | `finance@contoso.com` | ✅ Match (alias) |
| `Sales-Team` | `sales@contoso.com` | ❌ No match (wrong prefix) |

---

## Monitoring and Troubleshooting

### View Runbook Execution History

1. Navigate to your runbook
2. Click **Jobs** to see all executions
3. Click on a job to view:
   - **Output**: Script execution logs
   - **Errors**: Any errors encountered
   - **Warnings**: Non-critical issues

### Common Issues and Solutions

#### Issue: "Failed to connect to Exchange Online"

**Solution**: Verify Managed Identity has Exchange Administrator role
```powershell
# Check role assignment
Connect-MgGraph
$ManagedIdentityObjectId = "YOUR-OBJECT-ID"
Get-MgDirectoryRoleMember -DirectoryRoleId (Get-MgDirectoryRole -Filter "DisplayName eq 'Exchange Administrator'").Id |
    Where-Object {$_.Id -eq $ManagedIdentityObjectId}
```

#### Issue: "No groups found with prefix 'SM-'"

**Solution**: Verify groups exist and naming is correct
```powershell
Connect-ExchangeOnline
Get-DistributionGroup -Filter "Name -like 'SMBX-*'"
```

#### Issue: "Shared mailbox 'xyz' not found"

**Solution**: Ensure shared mailbox name matches group name (minus prefix)
```powershell
# List all shared mailboxes
Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table Name, Alias, PrimarySmtpAddress
```

#### Issue: Permissions not applying

**Solution**: Check if user is actually in the group
```powershell
Get-DistributionGroupMember -Identity "SM-sales@contoso.com"
```

---

## Advanced Configuration

### Custom Group Prefix

To use a different prefix (e.g., `SHARED-`):

1. Edit the runbook schedule parameters
2. Change **GROUPPREFIX** to `SHARED-`
3. Rename your groups accordingly

### Organization Parameter

The **ORGANIZATION** parameter is required and must be set to your tenant domain:
- Example: `contoso.onmicrosoft.com`
- This is required for Managed Identity authentication

### Multiple Runbooks for Different Prefixes

You can run multiple instances with different prefixes:

1. Create separate runbooks or schedules
2. Use different prefixes: `SMBX-`, `ROOM-`, `RESOURCE-`
3. Manage different mailbox types independently

### Email Notifications

Add email notifications for failures:

1. Navigate to **Automation Account** → **Alerts**
2. Click **+ New alert rule**
3. Configure conditions (e.g., Job failed)
4. Add action group with email notification

---

## Security Best Practices

✅ **Use Managed Identity** - No credentials to manage or rotate  
✅ **Least Privilege** - Only Exchange Administrator role assigned  
✅ **Audit Logs** - All changes logged in Exchange audit logs  
✅ **Regular Reviews** - Monitor runbook execution logs  
✅ **Test Environment** - Test changes in non-production first  

---

## Maintenance

### Monthly Tasks

- Review runbook execution history
- Check for failed jobs
- Verify group/mailbox naming consistency
- Update ExchangeOnlineManagement module if needed

### Updating the Script

1. Edit the runbook
2. Make changes
3. Save and test
4. Publish when ready

---

## Cost Considerations

**Azure Automation Pricing** (as of 2025):
- **Job runtime**: First 500 minutes/month free, then ~$0.002/minute
- **Watchers**: Not used in this solution
- **Estimated monthly cost**: $0-5 (depending on execution frequency and duration)

**Typical execution time**: 1-5 minutes depending on number of groups

---

## Support and Resources

- **Exchange Online PowerShell**: [Microsoft Docs](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
- **Azure Automation**: [Microsoft Docs](https://learn.microsoft.com/en-us/azure/automation/)
- **Managed Identity**: [Microsoft Docs](https://learn.microsoft.com/en-us/azure/automation/enable-managed-identity-for-automation)

---

## Quick Reference Commands

### Manual Execution (Local Testing)

```powershell
# Run locally with Managed Identity (requires Azure VM with MI)
.\Sync-SharedMailboxPermissions.ps1 -GroupPrefix "SMBX-" -Organization "contoso.onmicrosoft.com"
```

### Check Current Permissions

```powershell
Connect-ExchangeOnline

# List all shared mailboxes
Get-Mailbox -RecipientTypeDetails SharedMailbox | Format-Table Name, PrimarySmtpAddress

# Check specific mailbox permissions
Get-MailboxPermission -Identity "sales@contoso.com" | 
    Where-Object {$_.IsInherited -eq $false} | 
    Format-Table User, AccessRights, AutoMapping

# Check SendAs permissions
Get-RecipientPermission -Identity "sales@contoso.com" | 
    Format-Table Trustee, AccessRights
```

### List All SMBX- Groups

```powershell
Get-DistributionGroup -Filter "Name -like 'SMBX-*'" | 
    Format-Table Name, PrimarySmtpAddress, @{Name="Members";Expression={(Get-DistributionGroupMember -Identity $_.Identity).Count}}
```

---

## Next Steps

1. ✅ Complete the setup steps above
2. ✅ Test with a single shared mailbox
3. ✅ Create groups for all your shared mailboxes
4. ✅ Schedule the runbook
5. ✅ Monitor first few executions
6. ✅ Document your naming conventions for your team

---

**Setup Complete!** Your shared mailbox permissions will now be automatically managed based on group membership. 🎉
