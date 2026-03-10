<#
.SYNOPSIS
    Inventory all mail-enabled objects in Exchange Online
.DESCRIPTION
    Exports all mail-enabled recipients from Exchange Online to CSV for comparison with Exchange 2016
    This script helps assess readiness for enabling edge blocking in Exchange Online
.PARAMETER ExportPath
    Path where the CSV export will be saved. Default: C:\Temp\ExchangeOnline-MailEnabledObjects.csv
.EXAMPLE
    .\Get-ExchangeOnlineMailObjects.ps1
    .\Get-ExchangeOnlineMailObjects.ps1 -ExportPath "C:\Reports\Cloud-Recipients.csv"
.NOTES
    Author: Exchange Hybrid Assessment
    Version: 1.0
    Requires Exchange Online PowerShell module v3.0 or later
    Install: Install-Module -Name ExchangeOnlineManagement
#>

[CmdletBinding()]
param(
    [string]$ExportPath = "C:\Temp\ExchangeOnline-MailEnabledObjects.csv"
)

$ErrorActionPreference = "Stop"

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Exchange Online Mail-Enabled Objects Inventory" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    if (-not (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue)) {
        Write-Host "Exchange Online PowerShell module not loaded." -ForegroundColor Yellow
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false
    }

    $allRecipients = @()
    
    Write-Host "[1/7] Collecting Mailboxes..." -ForegroundColor Yellow
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties RecipientTypeDetails,EmailAddresses,ArchiveStatus,PrimarySmtpAddress | Select-Object `
        @{N='RecipientType';E={'UserMailbox'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'Cloud'}},
        @{N='ServerName';E={'Cloud'}},
        ArchiveStatus,
        @{N='IsRemote';E={$_.RecipientTypeDetails -like '*Remote*'}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $mailboxes
    Write-Host "      Found: $($mailboxes.Count) mailboxes" -ForegroundColor Green
    
    Write-Host "[2/7] Collecting Distribution Groups..." -ForegroundColor Yellow
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
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $distributionGroups
    Write-Host "      Found: $($distributionGroups.Count) distribution groups" -ForegroundColor Green
    
    Write-Host "[3/7] Collecting Mail-Enabled Security Groups..." -ForegroundColor Yellow
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
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $mailSecurityGroups
    Write-Host "      Found: $($mailSecurityGroups.Count) mail-enabled security groups" -ForegroundColor Green
    
    Write-Host "[4/7] Collecting Dynamic Distribution Groups..." -ForegroundColor Yellow
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
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $dynamicGroups
    Write-Host "      Found: $($dynamicGroups.Count) dynamic distribution groups" -ForegroundColor Green
    
    Write-Host "[5/7] Collecting Mail Contacts..." -ForegroundColor Yellow
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
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $mailContacts
    Write-Host "      Found: $($mailContacts.Count) mail contacts" -ForegroundColor Green
    
    Write-Host "[6/7] Collecting Mail Users..." -ForegroundColor Yellow
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
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={if($_.DistinguishedName){($_.DistinguishedName -replace '^CN=.+?,((?:OU|CN)=.+)','$1') -replace ',DC=.*$',''}else{'N/A - Cloud Only'}}},
        DistinguishedName
    
    $allRecipients += $mailUsers
    Write-Host "      Found: $($mailUsers.Count) mail users" -ForegroundColor Green
    
    Write-Host "[7/7] Collecting Microsoft 365 Groups..." -ForegroundColor Yellow
    $m365Groups = Get-UnifiedGroup -ResultSize Unlimited | Select-Object `
        @{N='RecipientType';E={'Microsoft365Group'}},
        DisplayName,
        PrimarySmtpAddress,
        Alias,
        RecipientTypeDetails,
        @{N='EmailAddresses';E={($_.EmailAddresses | Where-Object {$_ -like "smtp:*"}) -join ';'}},
        @{N='Database';E={'N/A'}},
        @{N='ServerName';E={'N/A'}},
        @{N='ArchiveStatus';E={'N/A'}},
        @{N='IsRemote';E={$false}},
        @{N='OrganizationalUnit';E={'N/A - Cloud Only'}},
        DistinguishedName
    
    $allRecipients += $m365Groups
    Write-Host "      Found: $($m365Groups.Count) Microsoft 365 groups" -ForegroundColor Green
    
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
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Data collection complete!" -ForegroundColor Green
    Write-Host "Next step: Run Compare-EdgeBlockingReadiness.ps1" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
