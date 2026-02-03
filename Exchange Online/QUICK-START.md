# Quick Start Guide - Shared Mailbox Permission Management

## What This Does

Automatically manages shared mailbox permissions based on group membership:
- **Group naming**: `SMBX-<mailboxname>` controls access to shared mailbox `<mailboxname>`
- **Auto-add**: Users added to group → get FullAccess + SendAs + Automapping
- **Auto-remove**: Users removed from group → lose all permissions
- **Runs automatically**: Scheduled in Azure Automation with Managed Identity

---

## 5-Minute Setup

### 1. Create Azure Automation Account

```powershell
# In Azure Portal:
# Create Resource → Automation → Create
# - Enable System Managed Identity ✓
# - Note the Object ID
```

### 2. Grant Exchange Admin Role

```powershell
# Azure AD → Roles and administrators → Exchange Administrator
# Add assignments → Select your Automation Account
```

### 3. Install Module

```powershell
# Automation Account → Modules → Add from gallery
# Search: ExchangeOnlineManagement → Import
```

### 4. Create Runbook

```powershell
# Automation Account → Runbooks → Create
# - Name: Sync-SharedMailboxPermissions
# - Type: PowerShell
# - Runtime: 7.2
# - Paste script content → Save → Publish
```

### 5. Schedule It

```powershell
# Runbook → Schedules → Add schedule
# - Daily at 2:00 AM
# - Parameters: GroupPrefix = "SMBX-", Organization = "contoso.onmicrosoft.com"
```

---

## Usage

### Create Access Group for Shared Mailbox

```powershell
# Example: Managing sales@contoso.com

# 1. Create group (if not exists)
New-DistributionGroup -Name "SMBX-sales@contoso.com" -Type Security -ManagedBy "admin@contoso.com"

# 2. Add users
Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "john@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "jane@contoso.com"

# 3. Wait for scheduled run (or trigger manually)
# Done! Users will have access with automapping.
```

### Remove Access

```powershell
# Remove user from group
Remove-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "john@contoso.com" -Confirm:$false

# Next sync run removes permissions automatically
```

---

## Naming Examples

| Group Name | Manages Mailbox | Status |
|-----------|----------------|--------|
| `SMBX-sales@contoso.com` | `sales@contoso.com` | ✅ Works |
| `SMBX-support@contoso.com` | `support@contoso.com` | ✅ Works |
| `SMBX-hr` | `hr@contoso.com` | ✅ Works |
| `Sales-Team` | `sales@contoso.com` | ❌ Wrong prefix |

---

## Verify It Works

```powershell
Connect-ExchangeOnline

# Check permissions
Get-MailboxPermission -Identity "sales@contoso.com" | 
    Where-Object {$_.IsInherited -eq $false} | 
    Format-Table User, AccessRights, AutoMapping

# Check SendAs
Get-RecipientPermission -Identity "sales@contoso.com" | 
    Format-Table Trustee, AccessRights
```

---

## Troubleshooting

### Script not running?
- Check Managed Identity has Exchange Administrator role
- Verify ExchangeOnlineManagement module installed
- Check runbook job history for errors

### Permissions not applying?
- Verify group name matches: `SMBX-<mailboxname>`
- Check user is actually in the group
- Ensure shared mailbox exists
- Wait for next scheduled run or trigger manually

### Automapping not working?
- User must restart Outlook
- Wait 5-10 minutes after permission grant
- Verify AutoMapping = True in permissions

---

## Files in This Solution

- **`Sync-SharedMailboxPermissions.ps1`** - Main script
- **`README-ManagedIdentity-Setup.md`** - Detailed setup guide
- **`EXAMPLES-Testing.md`** - Testing procedures and examples
- **`QUICK-START.md`** - This file

---

## Support

For detailed instructions, see `README-ManagedIdentity-Setup.md`

For testing procedures, see `EXAMPLES-Testing.md`

---

**That's it!** Your shared mailbox permissions are now automated. 🎉
