<#
.SYNOPSIS
    Compare Exchange 2016 and Exchange Online mail-enabled objects
.DESCRIPTION
    Identifies objects that exist in Exchange 2016 but are missing in Exchange Online
    Generates a comprehensive HTML report to determine if edge blocking can be safely enabled
.PARAMETER OnPremCsvPath
    Path to the Exchange 2016 CSV export. Default: C:\Temp\Exchange2016-MailEnabledObjects.csv
.PARAMETER CloudCsvPath
    Path to the Exchange Online CSV export. Default: C:\Temp\ExchangeOnline-MailEnabledObjects.csv
.PARAMETER ReportPath
    Path where the HTML report will be saved. Default: C:\Temp\EdgeBlocking-ReadinessReport.html
.EXAMPLE
    .\Compare-EdgeBlockingReadiness.ps1
    .\Compare-EdgeBlockingReadiness.ps1 -ReportPath "C:\Reports\EdgeBlocking-Assessment.html"
.NOTES
    Author: Exchange Hybrid Assessment
    Version: 1.0
    Run after collecting data from both Exchange 2016 and Exchange Online
#>

[CmdletBinding()]
param(
    [string]$OnPremCsvPath = "C:\Temp\Exchange2016-MailEnabledObjects.csv",
    [string]$CloudCsvPath = "C:\Temp\ExchangeOnline-MailEnabledObjects.csv",
    [string]$ReportPath = "C:\Temp\EdgeBlocking-ReadinessReport.html"
)

$ErrorActionPreference = "Stop"

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Edge Blocking Readiness Assessment" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Loading data files..." -ForegroundColor Yellow
    
    if (-not (Test-Path $OnPremCsvPath)) {
        throw "On-premises CSV not found: $OnPremCsvPath`nPlease run Get-Exchange2016MailObjects.ps1 first."
    }
    
    if (-not (Test-Path $CloudCsvPath)) {
        throw "Cloud CSV not found: $CloudCsvPath`nPlease run Get-ExchangeOnlineMailObjects.ps1 first."
    }
    
    $onPremRecipients = Import-Csv $OnPremCsvPath
    $cloudRecipients = Import-Csv $CloudCsvPath
    
    Write-Host "  On-Premises: $($onPremRecipients.Count) objects" -ForegroundColor Green
    Write-Host "  Cloud: $($cloudRecipients.Count) objects" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Building cloud recipient index..." -ForegroundColor Yellow
    $cloudEmails = @{}
    foreach ($recipient in $cloudRecipients) {
        if ($recipient.PrimarySmtpAddress) {
            $cloudEmails[$recipient.PrimarySmtpAddress.ToLower()] = $recipient
        }
    }
    
    Write-Host "Analyzing differences..." -ForegroundColor Yellow
    
    $missingInCloud = @()
    $onPremMailboxesNotInCloud = @()
    $publicFolders = @()
    $syncedObjects = @()
    
    foreach ($onPremRecipient in $onPremRecipients) {
        $email = $onPremRecipient.PrimarySmtpAddress.ToLower()
        
        if ($onPremRecipient.RecipientType -eq 'MailPublicFolder') {
            $publicFolders += $onPremRecipient
        }
        
        if ($cloudEmails.ContainsKey($email)) {
            $syncedObjects += $onPremRecipient
        } elseif ($onPremRecipient.RecipientType -ne 'RemoteMailbox') {
            $missingInCloud += $onPremRecipient
            
            if ($onPremRecipient.RecipientType -eq 'UserMailbox') {
                $onPremMailboxesNotInCloud += $onPremRecipient
            }
        }
    }
    
    Write-Host ""
    Write-Host "Analysis complete!" -ForegroundColor Green
    Write-Host "  Synced to cloud: $($syncedObjects.Count)" -ForegroundColor Green
    Write-Host "  Missing in cloud: $($missingInCloud.Count)" -ForegroundColor $(if($missingInCloud.Count -gt 0){'Red'}else{'Green'})
    Write-Host "  On-prem mailboxes NOT in cloud: $($onPremMailboxesNotInCloud.Count)" -ForegroundColor $(if($onPremMailboxesNotInCloud.Count -gt 0){'Red'}else{'Green'})
    Write-Host ""
    
    $canEnableEdgeBlocking = $true
    $blockers = @()
    $warnings = @()
    
    if ($onPremMailboxesNotInCloud.Count -gt 0) {
        $canEnableEdgeBlocking = $false
        $blockers += "Found $($onPremMailboxesNotInCloud.Count) on-premises mailboxes that do NOT exist in Exchange Online (mail would fail)"
    }
    
    if ($publicFolders.Count -gt 0) {
        $warnings += "Found $($publicFolders.Count) mail-enabled public folders - these need to be synchronized as MailUser objects in Exchange Online"
        
        $pfSyncedCount = 0
        foreach ($pf in $publicFolders) {
            $pfEmail = $pf.PrimarySmtpAddress.ToLower()
            if ($cloudEmails.ContainsKey($pfEmail)) {
                $pfSyncedCount++
            }
        }
        
        if ($pfSyncedCount -lt $publicFolders.Count) {
            $canEnableEdgeBlocking = $false
            $blockers += "Only $pfSyncedCount of $($publicFolders.Count) public folders are synchronized to Exchange Online"
        } else {
            $warnings += "All public folders are synchronized to Exchange Online - edge blocking is possible"
        }
    }
    
    if ($missingInCloud.Count -gt 0) {
        $nonMailboxMissing = $missingInCloud | Where-Object {$_.RecipientType -ne 'UserMailbox' -and $_.RecipientType -ne 'MailPublicFolder'}
        if ($nonMailboxMissing.Count -gt 0) {
            $canEnableEdgeBlocking = $false
            $blockers += "Found $($nonMailboxMissing.Count) mail-enabled objects (groups, contacts, mail users) missing in Exchange Online"
        }
    }
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    
    $htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Edge Blocking Readiness Report</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #106ebe; margin-top: 30px; border-bottom: 2px solid #e1e1e1; padding-bottom: 5px; }
        h3 { color: #323130; margin-top: 20px; }
        .summary { background-color: white; padding: 20px; border-radius: 5px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .status-ok { color: #107c10; font-weight: bold; font-size: 1.3em; }
        .status-warning { color: #ff8c00; font-weight: bold; font-size: 1.3em; }
        .status-error { color: #d13438; font-weight: bold; font-size: 1.3em; }
        table { border-collapse: collapse; width: 100%; background-color: white; margin-top: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background-color: #0078d4; color: white; padding: 12px; text-align: left; font-weight: 600; }
        td { padding: 10px; border-bottom: 1px solid #e1e1e1; }
        tr:hover { background-color: #f0f0f0; }
        tr:last-child td { border-bottom: none; }
        .blocker { background-color: #fde7e9; padding: 15px; border-left: 4px solid #d13438; margin: 10px 0; border-radius: 3px; }
        .warning { background-color: #fff4ce; padding: 15px; border-left: 4px solid #ff8c00; margin: 10px 0; border-radius: 3px; }
        .info { background-color: #e7f3ff; padding: 15px; border-left: 4px solid #0078d4; margin: 10px 0; border-radius: 3px; }
        .recommendation { background-color: #dff6dd; padding: 15px; border-left: 4px solid #107c10; margin: 10px 0; border-radius: 3px; }
        ul { line-height: 1.8; }
        ol { line-height: 1.8; }
        code { background-color: #f3f2f1; padding: 2px 6px; border-radius: 3px; font-family: 'Consolas', monospace; }
        .metric-value { font-size: 1.1em; font-weight: bold; }
        .timestamp { color: #605e5c; font-size: 0.9em; }
    </style>
</head>
<body>
    <h1>🔒 Exchange Online Edge Blocking Readiness Report</h1>
    <p class="timestamp">Generated: $timestamp</p>
    
    <div class="summary">
        <h2>📊 Executive Summary</h2>
        <p><strong>Can Enable Edge Blocking:</strong> <span class="status-$(if($canEnableEdgeBlocking){'ok'}else{'error'})">$(if($canEnableEdgeBlocking){'✅ YES'}else{'❌ NO'})</span></p>
        
        $(if($blockers.Count -gt 0) {
            "<h3>🚫 Blockers (Must Fix Before Enabling Edge Blocking):</h3>"
            foreach($blocker in $blockers) {
                "<div class='blocker'><strong>❌ $blocker</strong></div>"
            }
        })
        
        $(if($warnings.Count -gt 0) {
            "<h3>⚠️ Warnings (Review Required):</h3>"
            foreach($warning in $warnings) {
                "<div class='warning'><strong>⚠️ $warning</strong></div>"
            }
        })
        
        $(if($canEnableEdgeBlocking) {
            "<div class='recommendation'>
                <h3>✅ Environment is Ready!</h3>
                <p>All mail-enabled objects are properly synchronized to Exchange Online. You can proceed with enabling edge blocking.</p>
            </div>"
        })
    </div>
    
    <div class="summary">
        <h2>📈 Statistics</h2>
        <table>
            <tr><th>Metric</th><th>Count</th><th>Status</th></tr>
            <tr>
                <td>Total On-Premises Recipients</td>
                <td class="metric-value">$($onPremRecipients.Count)</td>
                <td>ℹ️ Info</td>
            </tr>
            <tr>
                <td>Total Cloud Recipients</td>
                <td class="metric-value">$($cloudRecipients.Count)</td>
                <td>ℹ️ Info</td>
            </tr>
            <tr>
                <td>Objects Synced to Cloud</td>
                <td class="metric-value">$($syncedObjects.Count)</td>
                <td style="color: #107c10;">✅ Good</td>
            </tr>
            <tr>
                <td>On-Premises Mailboxes (Not in Cloud)</td>
                <td class="metric-value" style="color: $(if($onPremMailboxesNotInCloud.Count -gt 0){'#d13438'}else{'#107c10'})">$($onPremMailboxesNotInCloud.Count)</td>
                <td style="color: $(if($onPremMailboxesNotInCloud.Count -gt 0){'#d13438'}else{'#107c10'})">$(if($onPremMailboxesNotInCloud.Count -gt 0){'❌ Blocker'}else{'✅ Good'})</td>
            </tr>
            <tr>
                <td>Mail-Enabled Public Folders</td>
                <td class="metric-value" style="color: $(if($publicFolders.Count -gt 0){'#ff8c00'}else{'#107c10'})">$($publicFolders.Count)</td>
                <td style="color: $(if($publicFolders.Count -gt 0){'#ff8c00'}else{'#107c10'})">$(if($publicFolders.Count -gt 0){'⚠️ Review'}else{'✅ None'})</td>
            </tr>
            <tr>
                <td>Objects Missing in Cloud</td>
                <td class="metric-value" style="color: $(if($missingInCloud.Count -gt 0){'#d13438'}else{'#107c10'})">$($missingInCloud.Count)</td>
                <td style="color: $(if($missingInCloud.Count -gt 0){'#d13438'}else{'#107c10'})">$(if($missingInCloud.Count -gt 0){'❌ Blocker'}else{'✅ Good'})</td>
            </tr>
        </table>
    </div>
    
    $(if($onPremMailboxesNotInCloud.Count -gt 0) {
        "<div class='summary'>
        <h2>📬 On-Premises Mailboxes (Not in Exchange Online)</h2>
        <div class='blocker'>
            <strong>Action Required:</strong> These mailboxes do NOT exist in Exchange Online. Mail delivery would fail. They must be migrated or synchronized before enabling edge blocking.
        </div>
        <table>
            <tr><th>Display Name</th><th>Email Address</th><th>Type</th><th>Organizational Unit</th><th>Database</th><th>Server</th></tr>"
        foreach($mb in $onPremMailboxesNotInCloud) {
            "<tr><td>$($mb.DisplayName)</td><td>$($mb.PrimarySmtpAddress)</td><td>$($mb.RecipientTypeDetails)</td><td>$($mb.OrganizationalUnit)</td><td>$($mb.Database)</td><td>$($mb.ServerName)</td></tr>"
        }
        "</table>
        <div class='info'>
            <h3>Migration Options:</h3>
            <ul>
                <li><strong>Hybrid Migration:</strong> Use New-MoveRequest for seamless migration</li>
                <li><strong>Cutover Migration:</strong> For smaller deployments (&lt;150 mailboxes)</li>
                <li><strong>Staged Migration:</strong> For larger deployments</li>
            </ul>
        </div>
        </div>"
    })
    
    $(if($publicFolders.Count -gt 0) {
        "<div class='summary'>
        <h2>📁 Mail-Enabled Public Folders</h2>
        <div class='warning'>
            <strong>Public Folder Configuration Required:</strong> Mail-enabled public folders must be synchronized to Exchange Online as MailUser objects.
        </div>
        <table>
            <tr><th>Display Name</th><th>Email Address</th><th>Organizational Unit</th><th>Status in Cloud</th></tr>"
        foreach($pf in $publicFolders) {
            $pfEmail = $pf.PrimarySmtpAddress.ToLower()
            $inCloud = $cloudEmails.ContainsKey($pfEmail)
            "<tr><td>$($pf.DisplayName)</td><td>$($pf.PrimarySmtpAddress)</td><td>$($pf.OrganizationalUnit)</td><td style='color: $(if($inCloud){'#107c10'}else{'#d13438'})'>$(if($inCloud){'✅ Synced'}else{'❌ Missing'})</td></tr>"
        }
        "</table>
        <div class='info'>
            <h3>📋 Public Folder Synchronization Steps:</h3>
            <ol>
                <li><strong>Download the sync script:</strong> <code>https://aka.ms/SyncMailPublicFolders</code></li>
                <li><strong>Run the script:</strong> Creates MailUser objects in Exchange Online for each mail-enabled public folder</li>
                <li><strong>Configure centralized mail transport:</strong> Ensures mail to public folders routes correctly</li>
                <li><strong>Test mail flow:</strong> Send test emails to public folders and verify delivery</li>
                <li><strong>Monitor:</strong> Check for any NDRs or delivery issues</li>
            </ol>
            <p><strong>Important:</strong> Public folder content remains on-premises. Only mail routing is affected by edge blocking.</p>
        </div>
        <div class='recommendation'>
            <h3>✅ Good News About Public Folders:</h3>
            <p>You <strong>CAN</strong> enable edge blocking with on-premises public folders! The key is proper synchronization:</p>
            <ul>
                <li>Public folder content stays on-premises (no migration needed)</li>
                <li>MailUser objects in Exchange Online handle mail routing</li>
                <li>Mail flows to Exchange Online, then routes to on-premises public folders</li>
                <li>Users access public folders normally via Outlook</li>
            </ul>
        </div>
        </div>"
    })
    
    $(if($missingInCloud.Count -gt 0) {
        $nonMailboxMissing = $missingInCloud | Where-Object {$_.RecipientType -ne 'UserMailbox' -and $_.RecipientType -ne 'MailPublicFolder'}
        if ($nonMailboxMissing.Count -gt 0) {
            "<div class='summary'>
            <h2>❓ Objects Missing in Exchange Online</h2>
            <div class='blocker'>
                <strong>Action Required:</strong> These mail-enabled objects must be synchronized to Exchange Online.
            </div>
            <table>
                <tr><th>Display Name</th><th>Email Address</th><th>Type</th><th>Organizational Unit</th><th>Synced to Cloud</th></tr>"
            foreach($missing in $nonMailboxMissing) {
                "<tr><td>$($missing.DisplayName)</td><td>$($missing.PrimarySmtpAddress)</td><td>$($missing.RecipientType)</td><td>$($missing.OrganizationalUnit)</td><td>$($missing.SyncedToCloud)</td></tr>"
            }
            "</table>
            <div class='info'>
                <h3>Synchronization Methods:</h3>
                <ul>
                    <li><strong>Azure AD Connect:</strong> Ensure these objects are in sync scope</li>
                    <li><strong>Manual Creation:</strong> Create corresponding objects in Exchange Online</li>
                    <li><strong>Directory Sync:</strong> Verify sync filters and OU selection</li>
                </ul>
            </div>
            </div>"
        }
    })
    
    <div class="summary">
        <h2>📝 Recommendations & Next Steps</h2>
        
        $(if($canEnableEdgeBlocking) {
            "<div class='recommendation'>
                <h3>✅ Ready to Enable Edge Blocking</h3>
                <p>Your environment meets all requirements. Follow these steps to enable edge blocking:</p>
                <ol>
                    <li><strong>Backup current configuration:</strong> Document MX records and connector settings</li>
                    <li><strong>Update MX records:</strong> Point to Exchange Online (*.mail.protection.outlook.com)</li>
                    <li><strong>Configure inbound connector:</strong> If needed for specific scenarios</li>
                    <li><strong>Test mail flow:</strong> Send test emails to all recipient types</li>
                    <li><strong>Monitor for 48 hours:</strong> Watch for NDRs or delivery issues</li>
                    <li><strong>Update documentation:</strong> Record the change and new mail flow</li>
                </ol>
            </div>"
        } else {
            "<div class='blocker'>
                <h3>❌ Not Ready for Edge Blocking</h3>
                <p>Complete the following tasks before enabling edge blocking:</p>
                <ol>"
            if ($onPremMailboxesNotInCloud.Count -gt 0) {
                "<li><strong>Migrate on-premises mailboxes:</strong> Move all $($onPremMailboxesNotInCloud.Count) mailboxes to Exchange Online or ensure they exist as mail-enabled objects in Exchange Online</li>"
            }
            if ($publicFolders.Count -gt 0) {
                $pfNotSynced = 0
                foreach ($pf in $publicFolders) {
                    if (-not $cloudEmails.ContainsKey($pf.PrimarySmtpAddress.ToLower())) {
                        $pfNotSynced++
                    }
                }
                if ($pfNotSynced -gt 0) {
                    "<li><strong>Synchronize public folders:</strong> Run Sync-MailPublicFolders.ps1 for $pfNotSynced public folders</li>"
                }
            }
            $nonMailboxMissing = $missingInCloud | Where-Object {$_.RecipientType -ne 'UserMailbox' -and $_.RecipientType -ne 'MailPublicFolder'}
            if ($nonMailboxMissing.Count -gt 0) {
                "<li><strong>Synchronize mail-enabled objects:</strong> Ensure $($nonMailboxMissing.Count) groups/contacts/mail users are in Exchange Online</li>"
            }
            "</ol>
            </div>"
        })
        
        <div class='info'>
            <h3>🔍 What is Edge Blocking?</h3>
            <p><strong>Edge blocking</strong> means all inbound mail flows directly to Exchange Online, bypassing on-premises Exchange servers. Benefits include:</p>
            <ul>
                <li><strong>Better Protection:</strong> Microsoft's advanced threat protection filters all inbound mail</li>
                <li><strong>Improved Performance:</strong> Direct delivery to cloud mailboxes</li>
                <li><strong>Reduced Complexity:</strong> Simplified mail flow architecture</li>
                <li><strong>Lower Costs:</strong> Reduced on-premises infrastructure requirements</li>
            </ul>
            <p><strong>Requirements:</strong></p>
            <ul>
                <li>All mail-enabled recipients must exist in Exchange Online (as mailboxes or mail users)</li>
                <li>MX records point to Exchange Online</li>
                <li>Outbound mail can still flow through on-premises if needed</li>
            </ul>
        </div>
    </div>
    
    <div class="summary">
        <h2>📚 Additional Resources</h2>
        <ul>
            <li><a href="https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/manage-mail-flow-using-third-party-cloud" target="_blank">Manage mail flow using a third-party cloud service with Exchange Online</a></li>
            <li><a href="https://learn.microsoft.com/en-us/exchange/hybrid-deployment/deploy-hybrid" target="_blank">Exchange Server Hybrid Deployments</a></li>
            <li><a href="https://learn.microsoft.com/en-us/exchange/collaboration-exo/public-folders/batch-migration-of-legacy-public-folders" target="_blank">Batch migration of public folders to Exchange Online</a></li>
            <li><a href="https://learn.microsoft.com/en-us/exchange/collaboration-exo/public-folders/sync-mail-public-folders" target="_blank">Synchronize mail-enabled public folders</a></li>
            <li><a href="https://learn.microsoft.com/en-us/exchange/mailbox-migration/mailbox-migration" target="_blank">Mailbox migrations in Exchange Online</a></li>
            <li><a href="https://aka.ms/SyncMailPublicFolders" target="_blank">Sync-MailPublicFolders.ps1 Script</a></li>
        </ul>
    </div>
    
    <div class="summary">
        <h2>📞 Support Information</h2>
        <p>If you need assistance with:</p>
        <ul>
            <li><strong>Mailbox migrations:</strong> Review Microsoft's migration guides or contact Microsoft Support</li>
            <li><strong>Public folder synchronization:</strong> Use the Sync-MailPublicFolders.ps1 script from Microsoft</li>
            <li><strong>Azure AD Connect issues:</strong> Check sync errors in Azure AD Connect Health</li>
            <li><strong>Mail flow testing:</strong> Use Exchange Online message trace and mail flow troubleshooter</li>
        </ul>
    </div>
    
    <hr style="margin-top: 40px; border: none; border-top: 2px solid #e1e1e1;">
    <p style="text-align: center; color: #605e5c; font-size: 0.9em;">
        Report generated by Edge Blocking Readiness Assessment Tool<br>
        $timestamp
    </p>
</body>
</html>
"@
    
    $reportDir = Split-Path $ReportPath -Parent
    if (-not (Test-Path $reportDir)) {
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    }
    
    $htmlReport | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Host "Report generated: $ReportPath" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Exporting detailed CSV files..." -ForegroundColor Cyan
    $csvDir = Split-Path $ReportPath -Parent
    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($ReportPath)
    
    if ($onPremMailboxesNotInCloud.Count -gt 0) {
        $csvPath = Join-Path $csvDir "$baseFileName-OnPremMailboxesNotInCloud.csv"
        $onPremMailboxesNotInCloud | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported: $csvPath" -ForegroundColor Green
    }
    
    if ($publicFolders.Count -gt 0) {
        $csvPath = Join-Path $csvDir "$baseFileName-PublicFolders.csv"
        $publicFoldersWithStatus = $publicFolders | Select-Object * -ExcludeProperty SyncedToCloud | Select-Object *,@{N='SyncedToCloud';E={$cloudEmails.ContainsKey($_.PrimarySmtpAddress.ToLower())}}
        $publicFoldersWithStatus | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported: $csvPath" -ForegroundColor Green
    }
    
    if ($missingInCloud.Count -gt 0) {
        $csvPath = Join-Path $csvDir "$baseFileName-MissingInCloud.csv"
        $missingInCloud | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported: $csvPath" -ForegroundColor Green
    }
    
    if ($syncedObjects.Count -gt 0) {
        $csvPath = Join-Path $csvDir "$baseFileName-SyncedToCloud.csv"
        $syncedObjects | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "  Exported: $csvPath" -ForegroundColor Green
    }
    
    Write-Host ""
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "EDGE BLOCKING READINESS SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Can Enable Edge Blocking: " -NoNewline
    Write-Host $(if($canEnableEdgeBlocking){'✅ YES'}else{'❌ NO'}) -ForegroundColor $(if($canEnableEdgeBlocking){'Green'}else{'Red'})
    Write-Host ""
    Write-Host "Statistics:" -ForegroundColor Yellow
    Write-Host "  On-Premises Mailboxes NOT in Cloud: $($onPremMailboxesNotInCloud.Count)" -ForegroundColor $(if($onPremMailboxesNotInCloud.Count -gt 0){'Red'}else{'Green'})
    Write-Host "  Mail-Enabled Public Folders: $($publicFolders.Count)" -ForegroundColor $(if($publicFolders.Count -gt 0){'Yellow'}else{'Green'})
    Write-Host "  Objects Synced to Cloud: $($syncedObjects.Count)" -ForegroundColor Green
    Write-Host "  Objects Missing in Cloud: $($missingInCloud.Count)" -ForegroundColor $(if($missingInCloud.Count -gt 0){'Red'}else{'Green'})
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    
    if ($blockers.Count -gt 0) {
        Write-Host ""
        Write-Host "🚫 BLOCKERS (Must Fix):" -ForegroundColor Red
        foreach ($blocker in $blockers) {
            Write-Host "  ❌ $blocker" -ForegroundColor Red
        }
    }
    
    if ($warnings.Count -gt 0) {
        Write-Host ""
        Write-Host "⚠️  WARNINGS (Review):" -ForegroundColor Yellow
        foreach ($warning in $warnings) {
            Write-Host "  ⚠️  $warning" -ForegroundColor Yellow
        }
    }
    
    if ($canEnableEdgeBlocking) {
        Write-Host ""
        Write-Host "✅ Your environment is ready for edge blocking!" -ForegroundColor Green
        Write-Host "   Review the HTML report for detailed next steps." -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Opening report in browser..." -ForegroundColor Cyan
    Start-Process $ReportPath
    
} catch {
    Write-Error "Error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
