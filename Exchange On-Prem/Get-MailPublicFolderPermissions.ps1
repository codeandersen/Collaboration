<#
.SYNOPSIS
    Export all mail-enabled public folders with Send As and Send on Behalf permissions
.DESCRIPTION
    Retrieves all mail-enabled public folders from Exchange and exports their Send As and Send on Behalf permissions
    with user details (email address or UPN)
.PARAMETER ExportPath
    Path where the CSV export will be saved. Default: C:\Temp\MailPublicFolder-Permissions.csv
.PARAMETER Environment
    Specify 'OnPremises' or 'Online' depending on your Exchange environment. Default: OnPremises
.EXAMPLE
    .\Get-MailPublicFolderPermissions.ps1
    .\Get-MailPublicFolderPermissions.ps1 -Environment Online
    .\Get-MailPublicFolderPermissions.ps1 -ExportPath "C:\Reports\PF-Permissions.csv" -Environment Online
.NOTES
    Author: Exchange Hybrid Assessment
    Version: 1.0
    For Exchange Online: Requires Exchange Online PowerShell module and active connection
    For On-Premises: Run from Exchange Management Shell
.DISCLAIMER
    This script is provided AS-IS, with no warranty - Use at own risk.
#>

[CmdletBinding()]
param(
    [string]$ExportPath = "C:\Temp\MailPublicFolder-Permissions.csv",
    [ValidateSet('OnPremises','Online')]
    [string]$Environment = "OnPremises"
)

$ErrorActionPreference = "Stop"

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Mail-Enabled Public Folder Permissions Export" -ForegroundColor Cyan
    Write-Host "Environment: $Environment" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # Check environment
    if ($Environment -eq 'Online') {
        if (-not (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)) {
            Write-Host "Exchange Online PowerShell module not loaded." -ForegroundColor Yellow
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
            Connect-ExchangeOnline -ShowBanner:$false
        }
    } else {
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            throw "Exchange Management Shell not loaded. Please run this from Exchange Management Shell."
        }
    }

    Write-Host "Collecting mail-enabled public folders..." -ForegroundColor Yellow
    $publicFolders = Get-MailPublicFolder -ResultSize Unlimited
    Write-Host "Found: $($publicFolders.Count) mail-enabled public folders" -ForegroundColor Green
    Write-Host ""

    $results = @()
    $processedCount = 0

    foreach ($pf in $publicFolders) {
        $processedCount++
        Write-Host "[$processedCount/$($publicFolders.Count)] Processing: $($pf.DisplayName)" -ForegroundColor Yellow

        # Get Send As permissions
        try {
            if ($Environment -eq 'Online') {
                $sendAsPermissions = Get-RecipientPermission -Identity $pf.Identity -ErrorAction Stop | 
                    Where-Object { $_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.AccessRights -contains 'SendAs' }
            } else {
                # For on-premises, use Get-ADPermission with DistinguishedName
                $allPermissions = Get-ADPermission -Identity $pf.DistinguishedName -ErrorAction Stop
                $sendAsPermissions = $allPermissions | Where-Object { 
                    $_.User -notlike 'NT AUTHORITY\*' -and 
                    $_.User -ne 'S-1-5-10' -and
                    $_.Deny -eq $false -and
                    $_.ExtendedRights -and
                    ($_.ExtendedRights | Where-Object { $_.ToString() -match 'Send.*As' })
                }
            }
            
            foreach ($permission in $sendAsPermissions) {
                $userEmail = $null
                $userUPN = $null
                $trusteeName = if ($Environment -eq 'Online') { $permission.Trustee } else { $permission.User.ToString() }
                
                # Try to get user details
                try {
                    if ($Environment -eq 'Online') {
                        $user = Get-EXORecipient -Identity $trusteeName -ErrorAction SilentlyContinue
                    } else {
                        $user = Get-Recipient -Identity $trusteeName -ErrorAction SilentlyContinue
                    }
                    
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
                    PublicFolderName = $pf.DisplayName
                    PublicFolderEmail = $pf.PrimarySmtpAddress
                    PublicFolderAlias = $pf.Alias
                    PermissionType = 'SendAs'
                    TrusteeName = $trusteeName
                    TrusteeEmail = $userEmail
                    TrusteeUPN = $userUPN
                    AccessRights = if ($Environment -eq 'Online') { ($permission.AccessRights -join ';') } else { 'Send-As' }
                    IsInherited = $permission.IsInherited
                }
            }
        } catch {
            Write-Warning "  Failed to get Send As permissions: $_"
        }

        # Get Send on Behalf permissions
        try {
            if ($pf.GrantSendOnBehalfTo -and $pf.GrantSendOnBehalfTo.Count -gt 0) {
                foreach ($trustee in $pf.GrantSendOnBehalfTo) {
                    $userEmail = $null
                    $userUPN = $null
                    $trusteeName = $trustee.ToString()
                    
                    # Try to get user details
                    try {
                        if ($Environment -eq 'Online') {
                            $user = Get-EXORecipient -Identity $trusteeName -ErrorAction SilentlyContinue
                        } else {
                            $user = Get-Recipient -Identity $trusteeName -ErrorAction SilentlyContinue
                        }
                        
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
                        PublicFolderName = $pf.DisplayName
                        PublicFolderEmail = $pf.PrimarySmtpAddress
                        PublicFolderAlias = $pf.Alias
                        PermissionType = 'SendOnBehalf'
                        TrusteeName = $trustee
                        TrusteeEmail = $userEmail
                        TrusteeUPN = $userUPN
                        AccessRights = 'SendOnBehalf'
                        IsInherited = $false
                    }
                }
            }
        } catch {
            Write-Warning "  Failed to get Send on Behalf permissions: $_"
        }
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Export Summary" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Public Folders Processed: $processedCount" -ForegroundColor Green
    Write-Host "Total Permissions Found: $($results.Count)" -ForegroundColor Green
    Write-Host ""

    if ($results.Count -gt 0) {
        $exportDir = Split-Path $ExportPath -Parent
        if (-not (Test-Path $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }
        
        $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported to: $ExportPath" -ForegroundColor Green
        Write-Host ""
        
        Write-Host "Permissions by Type:" -ForegroundColor Cyan
        $results | Group-Object PermissionType | Format-Table Name, Count -AutoSize
        
        Write-Host ""
        Write-Host "Top Public Folders by Permission Count:" -ForegroundColor Cyan
        $results | Group-Object PublicFolderName | 
            Sort-Object Count -Descending | 
            Select-Object -First 10 Name, Count | 
            Format-Table -AutoSize
    } else {
        Write-Host "No Send As or Send on Behalf permissions found on any mail-enabled public folders." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Export complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green

} catch {
    Write-Error "Error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
