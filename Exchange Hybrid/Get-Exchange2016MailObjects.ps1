<#
.SYNOPSIS
    Inventory all mail-enabled objects in Exchange 2016 on-premises
.DESCRIPTION
    Exports all mail-enabled recipients from Exchange 2016 to CSV for comparison with Exchange Online
    This script helps assess readiness for enabling edge blocking in Exchange Online
.PARAMETER ExportPath
    Path where the CSV export will be saved. Default: C:\Temp\Exchange2016-MailEnabledObjects.csv
.EXAMPLE
    .\Get-Exchange2016MailObjects.ps1
    .\Get-Exchange2016MailObjects.ps1 -ExportPath "C:\Reports\OnPrem-Recipients.csv"
.NOTES
    Author: Exchange Hybrid Assessment
    Version: 1.0
    Run this script on Exchange 2016 server or with Exchange Management Shell loaded
#>

[CmdletBinding()]
param(
    [string]$ExportPath = "C:\Temp\Exchange2016-MailEnabledObjects.csv"
)

$ErrorActionPreference = "Stop"

try {
    if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "Exchange Management Shell not loaded. Please run this from Exchange Management Shell or load the Exchange snap-in."
    }

    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Exchange 2016 Mail-Enabled Objects Inventory" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    $allRecipients = @()
    
    Write-Host "[1/8] Collecting Mailboxes..." -ForegroundColor Yellow
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'UserMailbox'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        Database,
        ServerName,
        ArchiveStatus,
        @{N='SyncedToCloud';E={$_.RemoteRecipientType -ne $null}},
        RemoteRecipientType,
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $mailboxes
    Write-Host "      Found: $($mailboxes.Count) mailboxes" -ForegroundColor Green
    
    Write-Host "[2/8] Collecting Remote Mailboxes..." -ForegroundColor Yellow
    $remoteMailboxes = Get-RemoteMailbox -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'RemoteMailbox'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A - Cloud'}},
        @{N='ServerName';E={'N/A - Cloud'}},
        @{N='ArchiveStatus';E={$_.ArchiveStatus}},
        @{N='SyncedToCloud';E={$true}},
        RemoteRecipientType,
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $remoteMailboxes
    Write-Host "      Found: $($remoteMailboxes.Count) remote mailboxes" -ForegroundColor Green
    
    Write-Host "[3/8] Collecting Distribution Groups..." -ForegroundColor Yellow
    $distributionGroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'DistributionGroup'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$_.IsDirSynced}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $distributionGroups
    Write-Host "      Found: $($distributionGroups.Count) distribution groups" -ForegroundColor Green
    
    Write-Host "[4/8] Collecting Mail-Enabled Security Groups..." -ForegroundColor Yellow
    $mailSecurityGroups = Get-DistributionGroup -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup'} | Select-Object `
        @{N='RecipientType';E={'MailSecurityGroup'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$_.IsDirSynced}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $mailSecurityGroups
    Write-Host "      Found: $($mailSecurityGroups.Count) mail-enabled security groups" -ForegroundColor Green
    
    Write-Host "[5/8] Collecting Dynamic Distribution Groups..." -ForegroundColor Yellow
    $dynamicGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'DynamicDistributionGroup'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$_.IsDirSynced}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $dynamicGroups
    Write-Host "      Found: $($dynamicGroups.Count) dynamic distribution groups" -ForegroundColor Green
    
    Write-Host "[6/8] Collecting Mail Contacts..." -ForegroundColor Yellow
    $mailContacts = Get-MailContact -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'MailContact'}},
        DisplayName,
        @{N='PrimarySmtpAddress';E={$_.ExternalEmailAddress -replace 'SMTP:',''}},
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$_.IsDirSynced}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $mailContacts
    Write-Host "      Found: $($mailContacts.Count) mail contacts" -ForegroundColor Green
    
    Write-Host "[7/8] Collecting Mail Users..." -ForegroundColor Yellow
    $mailUsers = Get-MailUser -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'MailUser'}},
        DisplayName,
        @{N='PrimarySmtpAddress';E={$_.ExternalEmailAddress -replace 'SMTP:',''}},
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$_.IsDirSynced}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $mailUsers
    Write-Host "      Found: $($mailUsers.Count) mail users" -ForegroundColor Green
    
    Write-Host "[8/8] Collecting Mail-Enabled Public Folders..." -ForegroundColor Yellow
    $publicFolders = Get-MailPublicFolder -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'MailPublicFolder'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='SyncedToCloud';E={$false}},
        @{N='RemoteRecipientType';E={'N/A'}},
        @{N='OrganizationalUnit';E={($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}},
        DistinguishedName
    
    $allRecipients += $publicFolders
    Write-Host "      Found: $($publicFolders.Count) mail-enabled public folders" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Total mail-enabled objects: $($allRecipients.Count)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $exportDir = Split-Path $ExportPath -Parent
    if (-not (Test-Path $exportDir)) {
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
    }
    
    $allRecipients | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $ExportPath" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Summary by Type:" -ForegroundColor Cyan
    $allRecipients | Group-Object RecipientType | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "CRITICAL FINDINGS" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "On-Premises Mailboxes (not synced to cloud):" -ForegroundColor Yellow
    $onPremOnly = $mailboxes | Where-Object {$_.SyncedToCloud -eq $false}
    Write-Host "  Count: $($onPremOnly.Count)" -ForegroundColor $(if($onPremOnly.Count -gt 0){'Red'}else{'Green'})
    if ($onPremOnly.Count -gt 0) {
        Write-Host "  WARNING: These mailboxes must be migrated before enabling edge blocking!" -ForegroundColor Red
        $onPremOnly | Select-Object DisplayName, PrimarySmtpAddress, RecipientTypeDetails, Database | Format-Table -AutoSize
    }
    
    Write-Host ""
    Write-Host "Mail-Enabled Public Folders:" -ForegroundColor Yellow
    Write-Host "  Count: $($publicFolders.Count)" -ForegroundColor $(if($publicFolders.Count -gt 0){'Yellow'}else{'Green'})
    if ($publicFolders.Count -gt 0) {
        Write-Host "  NOTE: Public folders require synchronization to Exchange Online as MailUser objects!" -ForegroundColor Yellow
        Write-Host "  Use: https://aka.ms/SyncMailPublicFolders" -ForegroundColor Cyan
        $publicFolders | Select-Object DisplayName, PrimarySmtpAddress | Format-Table -AutoSize
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Data collection complete!" -ForegroundColor Green
    Write-Host "Next step: Run Get-ExchangeOnlineMailObjects.ps1" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
