Clear-Host

# Connect to Exchange Online
#$credentials = get-credential;
#Connect-ExchangeOnline -Credential $credentials
#$SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $credentials -Authentication "Basic" -AllowRedirection;
#Import-PSSession $SccSession

$mailboxes = @("ph@multi-wing.com")

#$monthsToKeep = 3

#$sourceDate = (Get-Date).AddMonths(-$monthsToKeep)

$searchName = "PH_DeleteItems_PurgeEmails"

$contentQuery = "folderid:B65ECEE16F5C684EA15669D00055FF1E003F8B85B2080000"

# Clean-up any old searches from failed runs of this script
if (Get-ComplianceSearch -Identity $searchName) {
    Write-Host "Cleaning up any old searches from failed runs of this script"

    try {
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false | Out-Null
    }
    catch {
        Write-Host "Clean-up of old script runs failed!" -ForegroundColor Red
        break
    }
}

Write-Host "Creating new search for deleted items for $mailboxes"

New-ComplianceSearch -Name $searchName -ContentMatchQuery $contentQuery -ExchangeLocation $mailboxes -AllowNotFoundExchangeLocationsEnabled $true | Out-Null
                                                                            
Start-ComplianceSearch -Identity $searchName | Out-Null

Write-Host "Searching..."

while ((Get-ComplianceSearch -Identity $searchName).Status -ne "Completed") {
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 2
}

$items = (Get-ComplianceSearch -Identity $searchName).Items

if ($items -gt 0) {
    $searchStatistics = Get-ComplianceSearch -Identity $searchName | Select-Object -Expand SearchStatistics | Convertfrom-JSON

    $sources = $searchStatistics.ExchangeBinding.Sources | Where-Object { $_.ContentItems -gt 0 }

    Write-Host ""
    Write-Host "Total Items found matching query:" $items 
    Write-Host ""
    Write-Host "Items found in the following mailboxes"
    Write-Host "--------------------------------------"

    foreach ($source in $sources) {
        Write-Host $source.Name "has" $source.ContentItems "items of size" $source.ContentSize
    }

    Write-Host ""

    $iterations = 0;
    
    $itemsProcessed = 0
    
    while ($itemsProcessed -lt $items) {
        $iterations++

        Write-Host "Deleting items iteration $($iterations)"

        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

        while ((Get-ComplianceSearchAction -Identity "$($searchName)_Purge").Status -ne "Completed") { 
            Start-Sleep -Seconds 2
        }

        $itemsProcessed = $itemsProcessed + 10
       
        # Remove the search action so we can recreate it
        Remove-ComplianceSearchAction -Identity "$($searchName)_Purge" -Confirm:$false  
    }
} else {
    Write-Host "No items found"
}

Write-Host ""
Write-Host "COMPLETED!"