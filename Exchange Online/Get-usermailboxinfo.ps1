<# 
REQUIRES:
- ExchangeOnlineManagement

WHAT IT DOES:
- Connects to Exchange Online
- Gets all UserMailbox + SharedMailbox
- Collects TotalItemSize
- Exports a simple CSV

OUTPUT:
- .\Mailbox_Simple_Report_<timestamp>.csv
#>

$ErrorActionPreference = 'Stop'

# Connect
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get only user + shared mailboxes
$mbxs = Get-ExoMailbox -ResultSize Unlimited `
        -RecipientTypeDetails UserMailbox,SharedMailbox `
        -Properties DisplayName,PrimarySmtpAddress,EmailAddresses,UserPrincipalName,RecipientTypeDetails

Write-Host ("Found {0} mailboxes (user + shared)" -f $mbxs.Count) -ForegroundColor Cyan

# Collect sizes with a progress bar
$report = @()
$total  = [math]::Max(1, $mbxs.Count)
$index  = 0

foreach ($m in $mbxs) {
    $index++
    $pct = [int](($index / $total) * 100)
    Write-Progress -Activity "Collecting mailbox sizes" -Status "$index of $total : ${m.PrimarySmtpAddress}" -PercentComplete $pct

    $sizeGB = $null
    try {
        $stat = Get-ExoMailboxStatistics -Identity $m.PrimarySmtpAddress
        if ($stat.TotalItemSize -and $stat.TotalItemSize.Value -and ($stat.TotalItemSize.Value | Get-Member -Name ToBytes -MemberType Method)) {
            $bytes = $stat.TotalItemSize.Value.ToBytes()
        } elseif ($stat.TotalItemSize -and $stat.TotalItemSize.ToString() -match '\(([\d,]+) bytes\)') {
            $bytes = [int64]($matches[1] -replace ',','')
        }
        if ($bytes) { $sizeGB = [Math]::Round($bytes / 1GB, 2) }
    } catch {
        Write-Warning "Stats failed for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
    }

    # Gather all SMTP addresses, primary + aliases
    $primaryEmail = $m.PrimarySmtpAddress
    $aliases = ($m.EmailAddresses | Where-Object { $_ -ne "SMTP:$primaryEmail" } | ForEach-Object {
        ($_ -split ':')[1]
    } | Sort-Object -Unique) -join ';'

    $report += [PSCustomObject]@{
        DisplayName     = $m.DisplayName
        PrimaryEmail    = $primaryEmail
        UPN             = $m.UserPrincipalName
        Aliases         = $aliases
        MailboxSizeGB   = $sizeGB
        MailboxType     = $m.RecipientTypeDetails
    }
}
Write-Progress -Activity "Collecting mailbox sizes" -Completed

# Export
$ts  = Get-Date -Format "yyyyMMdd_HHmmss"
$out = ".\Mailbox_Simple_Report_$ts.csv"
$report | Sort-Object MailboxType, DisplayName | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8

Write-Host "Done. Report saved to: $out" -ForegroundColor Green