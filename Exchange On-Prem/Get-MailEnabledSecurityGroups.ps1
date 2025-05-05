if (-not (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
Import-Module ActiveDirectory

# Define the base OU to search
$baseOU = "OU=xxxx Hosting,DC=xxxxcorp,DC=net"
Write-Host "Searching for mail-enabled security groups in $baseOU" -ForegroundColor Cyan

# Define the OU to exclude
$excludeOU = "OU=xxxxx Groups,OU=Groups,OU=xxxx Hosting,DC=xxxxcorp,DC=nett"
Write-Host "Excluding mail-enabled security groups from $excludeOU" -ForegroundColor Yellow

# Get all mail-enabled security groups directly using Get-Recipient
# This is more efficient than using Get-ADObject with filtering
$allMailEnabledSecurityGroups = Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup -OrganizationalUnit $baseOU

# Filter out groups from the excluded OU
$mailEnabledSecurityGroups = $allMailEnabledSecurityGroups | Where-Object {
    -not $_.DistinguishedName.EndsWith($excludeOU)
}

Write-Host "Found $($allMailEnabledSecurityGroups.Count) total mail-enabled security groups" -ForegroundColor Cyan
Write-Host "After excluding $($allMailEnabledSecurityGroups.Count - $mailEnabledSecurityGroups.Count) groups, processing $($mailEnabledSecurityGroups.Count) groups" -ForegroundColor Yellow

# Prepare results array for CSV export
$results = @()
# Prepare summary hashtable to count groups by OU
$ouSummary = @{}

$totalGroups = $mailEnabledSecurityGroups.Count
$currentGroup = 0

foreach ($group in $mailEnabledSecurityGroups) {
    $currentGroup++
    Write-Progress -Activity "Processing Mail-Enabled Security Groups" -Status "Processing $currentGroup of $totalGroups" -PercentComplete (($currentGroup / $totalGroups) * 100)
    
    # Get the parent OU for summary counting
    $parentOU = ($group.DistinguishedName -split ',', 2)[1]
    
    # Increment the count for this OU
    if ($ouSummary.ContainsKey($parentOU)) {
        $ouSummary[$parentOU]++
    } else {
        $ouSummary[$parentOU] = 1
    }
    
    # Get the primary SMTP address
    $primarySMTP = $group.PrimarySmtpAddress
    
    # Add to results
    $results += [PSCustomObject]@{
        DN = $group.DistinguishedName
        PrimarySMTP = $primarySMTP
        EmailAddresses = $group.EmailAddresses -join ";"
    }
}

# Generate timestamp for filenames
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Export the main results to CSV
$mainFilename = "MailEnabledSecurityGroups_$timestamp.csv"
$results | Export-Csv -Path $mainFilename -NoTypeInformation -Encoding Unicode

# Create summary report
$summaryFilename = "MailEnabledSecurityGroups_Summary_$timestamp.txt"
$summaryContent = "Mail-Enabled Security Groups Summary Report`r`n"
$summaryContent += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n"
$summaryContent += "Base OU: $baseOU`r`n"
$summaryContent += "Excluded OU: $excludeOU`r`n`r`n"
$summaryContent += "Count by Organizational Unit:`r`n"

foreach ($ou in $ouSummary.Keys | Sort-Object) {
    $summaryContent += "$ou : $($ouSummary[$ou])`r`n"
}

$summaryContent += "`r`nTotal Mail-Enabled Security Groups: $totalGroups`r`n"
$summaryContent | Out-File -FilePath $summaryFilename -Encoding Unicode

Write-Host "Script completed." -ForegroundColor Green
Write-Host "Main results exported to: $mainFilename" -ForegroundColor Green
Write-Host "Summary report exported to: $summaryFilename" -ForegroundColor Green
