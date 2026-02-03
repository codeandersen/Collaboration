# Shared Mailbox Permission Management - Examples & Testing

## Quick Start Examples

### Example 1: Sales Team Shared Mailbox

**Scenario**: You have a sales team that needs access to `sales@contoso.com`

```powershell
# 1. Create the shared mailbox (if not exists)
Connect-ExchangeOnline
New-Mailbox -Shared -Name "Sales Team" -PrimarySmtpAddress "sales@contoso.com"

# 2. Create the distribution group with SMBX- prefix
New-DistributionGroup -Name "SMBX-sales@contoso.com" -Type Security -ManagedBy "admin@contoso.com"

# 3. Add team members
Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "john.doe@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "jane.smith@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-sales@contoso.com" -Member "bob.wilson@contoso.com"

# 4. Run the sync script (or wait for scheduled run)
# The script will automatically grant FullAccess + SendAs with automapping

# 5. Verify permissions
Get-MailboxPermission -Identity "sales@contoso.com" | 
    Where-Object {$_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*"} |
    Format-Table User, AccessRights, AutoMapping
```

**Expected Result**: All three users will have FullAccess with automapping enabled and SendAs permissions.

---

### Example 2: Support Mailbox with User Rotation

**Scenario**: Support team members change frequently

```powershell
# Initial setup
New-Mailbox -Shared -Name "Support" -PrimarySmtpAddress "support@contoso.com"
New-DistributionGroup -Name "SMBX-support@contoso.com" -Type Security -ManagedBy "admin@contoso.com"

# Add initial support team
Add-DistributionGroupMember -Identity "SMBX-support@contoso.com" -Member "alice@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-support@contoso.com" -Member "bob@contoso.com"

# Run sync - both users get access
# ... time passes ...

# Alice leaves, Charlie joins
Remove-DistributionGroupMember -Identity "SMBX-support@contoso.com" -Member "alice@contoso.com" -Confirm:$false
Add-DistributionGroupMember -Identity "SMBX-support@contoso.com" -Member "charlie@contoso.com"

# Next sync run will:
# - Remove Alice's permissions
# - Add Charlie's permissions
# - Keep Bob's permissions unchanged
```

---

### Example 3: Multiple Shared Mailboxes

**Scenario**: Managing multiple departments

```powershell
# HR Mailbox
New-Mailbox -Shared -Name "HR" -PrimarySmtpAddress "hr@contoso.com"
New-DistributionGroup -Name "SMBX-hr@contoso.com" -Type Security -ManagedBy "admin@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-hr@contoso.com" -Member "hr.manager@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-hr@contoso.com" -Member "hr.assistant@contoso.com"

# Finance Mailbox
New-Mailbox -Shared -Name "Finance" -PrimarySmtpAddress "finance@contoso.com"
New-DistributionGroup -Name "SMBX-finance@contoso.com" -Type Security -ManagedBy "admin@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-finance@contoso.com" -Member "cfo@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-finance@contoso.com" -Member "accountant@contoso.com"

# IT Helpdesk Mailbox
New-Mailbox -Shared -Name "IT Helpdesk" -PrimarySmtpAddress "helpdesk@contoso.com"
New-DistributionGroup -Name "SMBX-helpdesk@contoso.com" -Type Security -ManagedBy "admin@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-helpdesk@contoso.com" -Member "it.admin@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-helpdesk@contoso.com" -Member "it.support@contoso.com"

# One script run manages all three mailboxes automatically
```

---

## Testing Procedures

### Pre-Deployment Testing

#### Test 1: Basic Functionality

```powershell
# 1. Create test mailbox
New-Mailbox -Shared -Name "Test Mailbox" -PrimarySmtpAddress "test-mailbox@contoso.com"

# 2. Create test group
New-DistributionGroup -Name "SMBX-test-mailbox@contoso.com" -Type Security -ManagedBy "admin@contoso.com"

# 3. Add yourself as test user
Add-DistributionGroupMember -Identity "SMBX-test-mailbox@contoso.com" -Member "your.email@contoso.com"

# 4. Run script manually (from Azure Automation or locally with cert auth)
# Check output for success message

# 5. Verify permissions
Get-MailboxPermission -Identity "test-mailbox@contoso.com" | 
    Where-Object {$_.User -eq "your.email@contoso.com"}

# Expected: FullAccess with AutoMapping = True

Get-RecipientPermission -Identity "test-mailbox@contoso.com" | 
    Where-Object {$_.Trustee -eq "your.email@contoso.com"}

# Expected: SendAs permission
```

#### Test 2: User Removal

```powershell
# 1. Remove yourself from the group
Remove-DistributionGroupMember -Identity "SMBX-test-mailbox@contoso.com" -Member "your.email@contoso.com" -Confirm:$false

# 2. Run script again

# 3. Verify permissions removed
Get-MailboxPermission -Identity "test-mailbox@contoso.com" | 
    Where-Object {$_.User -eq "your.email@contoso.com"}

# Expected: No results (permissions removed)
```

#### Test 3: Empty Group Handling

```powershell
# 1. Ensure group is empty
Get-DistributionGroupMember -Identity "SMBX-test-mailbox@contoso.com"

# 2. Manually add a permission
Add-MailboxPermission -Identity "test-mailbox@contoso.com" -User "someone@contoso.com" -AccessRights FullAccess

# 3. Run script

# 4. Verify all permissions removed
Get-MailboxPermission -Identity "test-mailbox@contoso.com" | 
    Where-Object {$_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*"}

# Expected: No results (all explicit permissions removed)
```

#### Test 4: Non-Existent Mailbox

```powershell
# 1. Create group for non-existent mailbox
New-DistributionGroup -Name "SMBX-nonexistent@contoso.com" -Type Security -ManagedBy "admin@contoso.com"
Add-DistributionGroupMember -Identity "SMBX-nonexistent@contoso.com" -Member "someone@contoso.com"

# 2. Run script

# 3. Check output
# Expected: Warning message "Shared mailbox 'nonexistent' not found. Skipping."
# Script should continue processing other groups
```

---

### Post-Deployment Validation

#### Validation Checklist

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# 1. List all SMBX- groups
$SMBXGroups = Get-DistributionGroup -Filter "Name -like 'SMBX-*'"
Write-Host "Found $($SMBXGroups.Count) SMBX- groups" -ForegroundColor Cyan
$SMBXGroups | Format-Table Name, PrimarySmtpAddress

# 2. For each group, verify corresponding mailbox exists
foreach ($group in $SMBXGroups) {
    $mailboxName = $group.Name -replace '^SMBX-', ''
    $mailbox = Get-Mailbox -Identity $mailboxName -ErrorAction SilentlyContinue
    
    if ($mailbox) {
        Write-Host "✓ $($group.Name) -> $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
    } else {
        Write-Host "✗ $($group.Name) -> Mailbox not found!" -ForegroundColor Red
    }
}

# 3. Verify permissions match group membership
foreach ($group in $SMBXGroups) {
    $mailboxName = $group.Name -replace '^SMBX-', ''
    $mailbox = Get-Mailbox -Identity $mailboxName -ErrorAction SilentlyContinue
    
    if ($mailbox) {
        Write-Host "`nChecking: $($group.Name)" -ForegroundColor Cyan
        
        # Get group members
        $members = Get-DistributionGroupMember -Identity $group.Identity | 
            Where-Object {$_.RecipientType -eq "UserMailbox"} |
            Select-Object -ExpandProperty PrimarySmtpAddress
        
        # Get mailbox permissions
        $permissions = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress | 
            Where-Object {$_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*"}
        
        Write-Host "  Group members: $($members.Count)" -ForegroundColor Gray
        Write-Host "  Mailbox delegates: $($permissions.Count)" -ForegroundColor Gray
        
        # Compare
        $comparison = Compare-Object -ReferenceObject $members -DifferenceObject ($permissions.User)
        
        if ($comparison) {
            Write-Host "  ⚠ Mismatch detected!" -ForegroundColor Yellow
            $comparison | Format-Table InputObject, SideIndicator
        } else {
            Write-Host "  ✓ In sync" -ForegroundColor Green
        }
    }
}
```

---

## User Experience Testing

### End User Perspective

#### Test: Automapping in Outlook

1. **Add user to SM- group**
2. **Run sync script**
3. **User actions**:
   - Close Outlook completely
   - Wait 5-10 minutes (for Exchange to process)
   - Open Outlook
   - **Expected**: Shared mailbox appears automatically in folder list
   - No manual "Add Shared Mailbox" needed

#### Test: SendAs Functionality

```powershell
# After permissions granted, user should be able to:
# 1. In Outlook, compose new email
# 2. Click "From" button
# 3. Select the shared mailbox address
# 4. Send email as the shared mailbox
```

#### Test: Permission Removal

1. **Remove user from SM- group**
2. **Run sync script**
3. **User actions**:
   - Close Outlook
   - Wait 5-10 minutes
   - Open Outlook
   - **Expected**: Shared mailbox no longer appears
   - User cannot access mailbox

---

## Performance Testing

### Benchmark Script

```powershell
# Measure execution time for different scales

# Small scale (1-5 groups)
Measure-Command {
    .\Sync-SharedMailboxPermissions.ps1 -AuthMethod ManagedIdentity
}

# Expected: 30-60 seconds

# Medium scale (10-20 groups)
# Expected: 1-3 minutes

# Large scale (50+ groups)
# Expected: 5-10 minutes
```

### Load Testing

```powershell
# Create multiple test groups and mailboxes
1..10 | ForEach-Object {
    $name = "test$_"
    New-Mailbox -Shared -Name "Test $name" -PrimarySmtpAddress "$name@contoso.com"
    New-DistributionGroup -Name "SMBX-$name@contoso.com" -Type Security -ManagedBy "admin@contoso.com"
    Add-DistributionGroupMember -Identity "SMBX-$name@contoso.com" -Member "testuser@contoso.com"
}

# Run script and monitor performance
Measure-Command {
    .\Sync-SharedMailboxPermissions.ps1 -Organization "contoso.onmicrosoft.com" -Verbose
}

# Cleanup
1..10 | ForEach-Object {
    $name = "test$_"
    Remove-Mailbox -Identity "$name@contoso.com" -Confirm:$false
    Remove-DistributionGroup -Identity "SMBX-$name@contoso.com" -Confirm:$false
}
```

---

## Troubleshooting Test Scenarios

### Scenario 1: Permission Not Applying

```powershell
# Debug steps
$mailbox = "sales@contoso.com"
$user = "john.doe@contoso.com"

# 1. Verify user is in group
Get-DistributionGroupMember -Identity "SMBX-sales@contoso.com" | 
    Where-Object {$_.PrimarySmtpAddress -eq $user}

# 2. Check if mailbox exists and is shared
Get-Mailbox -Identity $mailbox | 
    Select-Object Name, RecipientTypeDetails, PrimarySmtpAddress

# 3. Try manual permission grant
Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping $true

# 4. Check for errors
$Error[0] | Format-List * -Force
```

### Scenario 2: Automapping Not Working

```powershell
# Check AutoMapping setting
Get-MailboxPermission -Identity "sales@contoso.com" | 
    Where-Object {$_.User -eq "john.doe@contoso.com"} | 
    Select-Object User, AccessRights, AutoMapping, IsInherited

# If AutoMapping = False, fix it:
Remove-MailboxPermission -Identity "sales@contoso.com" -User "john.doe@contoso.com" -AccessRights FullAccess -Confirm:$false
Add-MailboxPermission -Identity "sales@contoso.com" -User "john.doe@contoso.com" -AccessRights FullAccess -AutoMapping $true

# User must restart Outlook for automapping to take effect
```

### Scenario 3: Script Runs But No Changes

```powershell
# Enable verbose output in Azure Automation
# Edit runbook parameters and add: -Verbose

# Or run locally with verbose
.\Sync-SharedMailboxPermissions.ps1 -AuthMethod ManagedIdentity -Verbose

# Check for:
# - "Permissions already in sync" messages
# - Group member count vs delegate count
# - Any warning messages
```

---

## Regression Testing

### After Script Updates

```powershell
# Test suite to run after any script modifications

# Test 1: Basic add permission
# Test 2: Basic remove permission
# Test 3: Multiple users
# Test 4: Empty group
# Test 5: Non-existent mailbox
# Test 6: Special characters in names
# Test 7: Large group (50+ members)
# Test 8: Concurrent modifications

# Create test report
$results = @()
$results += [PSCustomObject]@{Test="Add Permission"; Status="Pass"; Time="5s"}
$results += [PSCustomObject]@{Test="Remove Permission"; Status="Pass"; Time="4s"}
# ... etc

$results | Format-Table -AutoSize
```

---

## Monitoring and Alerts

### Set Up Test Alerts

```powershell
# In Azure Automation, create alert rules for:

# 1. Job Failure
# Condition: Job status = Failed
# Action: Email notification

# 2. Long Running Jobs
# Condition: Job duration > 10 minutes
# Action: Email notification

# 3. High Error Rate
# Condition: Error count > 5 in output
# Action: Email notification
```

### Health Check Script

```powershell
# Run weekly to verify system health

$report = @{
    Date = Get-Date
    TotalGroups = 0
    TotalMailboxes = 0
    OrphanedGroups = @()
    MismatchedPermissions = @()
}

$groups = Get-DistributionGroup -Filter "Name -like 'SMBX-*'"
$report.TotalGroups = $groups.Count

foreach ($group in $groups) {
    $mailboxName = $group.Name -replace '^SMBX-', ''
    $mailbox = Get-Mailbox -Identity $mailboxName -ErrorAction SilentlyContinue
    
    if (-not $mailbox) {
        $report.OrphanedGroups += $group.Name
    } else {
        $report.TotalMailboxes++
    }
}

# Output report
$report | ConvertTo-Json -Depth 3 | Out-File "HealthCheck-$(Get-Date -Format 'yyyyMMdd').json"
```

---

## Best Practices for Testing

✅ **Test in non-production first** - Use a test tenant or test mailboxes  
✅ **Document test results** - Keep a log of what was tested and outcomes  
✅ **Test edge cases** - Empty groups, special characters, large groups  
✅ **Verify end-user experience** - Test in actual Outlook client  
✅ **Monitor first production runs** - Watch closely for the first week  
✅ **Have rollback plan** - Know how to manually remove permissions if needed  

---

## Quick Rollback Procedure

If something goes wrong:

```powershell
# Emergency: Remove all permissions managed by a specific group
$group = "SMBX-sales@contoso.com"
$mailboxName = $group -replace '^SMBX-', ''
$members = Get-DistributionGroupMember -Identity $group

foreach ($member in $members) {
    Remove-MailboxPermission -Identity $mailboxName -User $member.PrimarySmtpAddress -AccessRights FullAccess -Confirm:$false
    Remove-RecipientPermission -Identity $mailboxName -Trustee $member.PrimarySmtpAddress -AccessRights SendAs -Confirm:$false
}

Write-Host "All permissions removed for $mailboxName" -ForegroundColor Yellow
```

---

## Testing Checklist

Before going to production:

- [ ] Tested with single mailbox and single user
- [ ] Tested adding multiple users
- [ ] Tested removing users
- [ ] Tested empty group scenario
- [ ] Tested non-existent mailbox scenario
- [ ] Verified automapping works in Outlook
- [ ] Verified SendAs permissions work
- [ ] Tested with 5+ mailboxes simultaneously
- [ ] Reviewed all script output for errors
- [ ] Documented naming conventions for team
- [ ] Created runbook schedule
- [ ] Set up monitoring alerts
- [ ] Tested manual runbook trigger
- [ ] Verified Managed Identity permissions
- [ ] Created rollback procedure document

---

**Ready for Production!** Once all tests pass, you can confidently deploy to production. 🚀
