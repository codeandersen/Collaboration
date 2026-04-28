<#
.SYNOPSIS
    Gets all shared mailboxes with Send As and Send On Behalf permissions in Exchange 2016.

.DESCRIPTION
    This script retrieves all shared mailboxes from Exchange 2016 and exports their Send As 
    and Send On Behalf permissions to a CSV file. Each permission is listed as a separate row
    with the shared mailbox email address and the permission type.

.PARAMETER OutputPath
    The path where the CSV file will be saved. Default is the current directory.

.EXAMPLE
    .\Get-SharedMailboxPermissions.ps1
    Exports shared mailbox permissions to SharedMailboxPermissions.csv in the current directory.

.EXAMPLE
    .\Get-SharedMailboxPermissions.ps1 -OutputPath "C:\Reports\Permissions.csv"
    Exports shared mailbox permissions to the specified path.

.NOTES
    Author: Exchange Admin
    Date: 2026-03-24
    Requires: Exchange Management Shell
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\SharedMailboxPermissions.csv"
)

$ErrorActionPreference = "Continue"

Write-Host "Starting shared mailbox permission collection..." -ForegroundColor Cyan

$results = @()

try {
    Write-Host "Retrieving all shared mailboxes..." -ForegroundColor Yellow
    $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    
    if ($sharedMailboxes.Count -eq 0) {
        Write-Host "No shared mailboxes found." -ForegroundColor Red
        return
    }
    
    Write-Host "Found $($sharedMailboxes.Count) shared mailbox(es). Processing permissions..." -ForegroundColor Green
    
    $counter = 0
    foreach ($mailbox in $sharedMailboxes) {
        $counter++
        Write-Progress -Activity "Processing Shared Mailboxes" -Status "Processing $($mailbox.DisplayName) ($counter of $($sharedMailboxes.Count))" -PercentComplete (($counter / $sharedMailboxes.Count) * 100)
        
        $mailboxEmail = $mailbox.PrimarySmtpAddress
        $mailboxName = $mailbox.DisplayName
        
        Write-Host "Processing: $mailboxName ($mailboxEmail)" -ForegroundColor Cyan
        
        # Get Send As permissions
        try {
            $sendAsPerms = Get-ADPermission -Identity $mailbox.DistinguishedName | 
                Where-Object { 
                    $_.ExtendedRights -like "*Send-As*" -and 
                    $_.IsInherited -eq $false -and 
                    $_.User -notlike "NT AUTHORITY\SELF" -and
                    $_.User -notlike "S-1-5-*"
                }
            
            foreach ($perm in $sendAsPerms) {
                $userEmail = $null
                $userUPN = $null
                $trusteeName = $perm.User.ToString()
                
                # Try to get user details
                try {
                    $user = Get-Recipient -Identity $trusteeName -ErrorAction SilentlyContinue
                    
                    if ($user) {
                        $userEmail = $user.PrimarySmtpAddress
                        if ($user.PSObject.Properties.Name -contains 'UserPrincipalName') {
                            $userUPN = $user.UserPrincipalName
                        }
                    }
                } catch {
                    Write-Verbose "Could not resolve user details for: $trusteeName"
                }
                
                $results += [PSCustomObject]@{
                    SharedMailboxName = $mailboxName
                    SharedMailboxEmail = $mailboxEmail
                    GrantedToName = $trusteeName
                    GrantedToEmail = $userEmail
                    GrantedToUPN = $userUPN
                    PermissionType = "Send As"
                    AccessRights = "Send As"
                }
                Write-Host "  - Send As: $trusteeName" -ForegroundColor Gray
                if ($userEmail) {
                    Write-Host "    Email: $userEmail" -ForegroundColor DarkGray
                }
            }
        }
        catch {
            Write-Host "  Error retrieving Send As permissions: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Get Send On Behalf permissions
        try {
            if ($mailbox.GrantSendOnBehalfTo -and $mailbox.GrantSendOnBehalfTo.Count -gt 0) {
                foreach ($delegate in $mailbox.GrantSendOnBehalfTo) {
                    $userEmail = $null
                    $userUPN = $null
                    $trusteeName = $delegate.ToString()
                    
                    # Try to get user details
                    try {
                        $user = Get-Recipient -Identity $trusteeName -ErrorAction SilentlyContinue
                        
                        if ($user) {
                            $userEmail = $user.PrimarySmtpAddress
                            if ($user.PSObject.Properties.Name -contains 'UserPrincipalName') {
                                $userUPN = $user.UserPrincipalName
                            }
                        }
                    } catch {
                        Write-Verbose "Could not resolve user details for: $trusteeName"
                    }
                    
                    $results += [PSCustomObject]@{
                        SharedMailboxName = $mailboxName
                        SharedMailboxEmail = $mailboxEmail
                        GrantedToName = $trusteeName
                        GrantedToEmail = $userEmail
                        GrantedToUPN = $userUPN
                        PermissionType = "Send On Behalf"
                        AccessRights = "Send On Behalf"
                    }
                    Write-Host "  - Send On Behalf: $trusteeName" -ForegroundColor Gray
                    if ($userEmail) {
                        Write-Host "    Email: $userEmail" -ForegroundColor DarkGray
                    }
                }
            }
        }
        catch {
            Write-Host "  Error retrieving Send On Behalf permissions: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Processing Shared Mailboxes" -Completed
    
    if ($results.Count -eq 0) {
        Write-Host "`nNo Send As or Send On Behalf permissions found on any shared mailboxes." -ForegroundColor Yellow
    }
    else {
        Write-Host "`nExporting $($results.Count) permission(s) to CSV..." -ForegroundColor Yellow
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Export completed successfully!" -ForegroundColor Green
        Write-Host "Output file: $OutputPath" -ForegroundColor Cyan
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "  Total Shared Mailboxes: $($sharedMailboxes.Count)" -ForegroundColor White
        Write-Host "  Total Permissions Found: $($results.Count)" -ForegroundColor White
        Write-Host "  Send As Permissions: $(($results | Where-Object {$_.PermissionType -eq 'Send As'}).Count)" -ForegroundColor White
        Write-Host "  Send On Behalf Permissions: $(($results | Where-Object {$_.PermissionType -eq 'Send On Behalf'}).Count)" -ForegroundColor White
    }
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
