#Requires -Modules ExchangeOnlineManagement, Az.Accounts

<#
.SYNOPSIS
    Sets retention policy on all mailboxes except those in the exempt security group (PARALLEL VERSION).

.DESCRIPTION
    This script connects to Exchange Online using managed identity,
    processes mailboxes using parallel runspaces for maximum performance,
    excludes members of the exempt security group,
    only enables archiving for users with valid SKUs (SPE_E5, SPE_E3, SPE_F5, ENTERPRISEPREMIUM, ENTERPRISEPACK, EXCHANGEENTERPRISE),
    and applies the specified retention policy in WhatIf mode.
    Optimized for large environments (25,000+ mailboxes).

.NOTES
    Designed for Azure Automation Account
    Requires ExchangeOnlineManagement module
    Runs in WhatIf mode by default
    Uses parallel processing with runspaces for 5-10x performance improvement
#>

param(
    [string]$ExemptSecurityGroup = "Exempt_OnlineArchive@starkworkspace.onmicrosoft.com",
    [string]$RetentionPolicyName = "STARK Group Default",
    [switch]$WhatIf = $true,
    [int]$ThrottleLimit = 10
)

write-output "Whatif is set to $WhatIf"
write-output "Parallel processing with $ThrottleLimit concurrent threads"

try {
    Write-Output "========================================"
    Write-Output "Retention Policy Assignment Script (PARALLEL)"
    Write-Output "========================================"
    Write-Output "Exempt Security Group: $ExemptSecurityGroup"
    Write-Output "Retention Policy: $RetentionPolicyName"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "Throttle Limit: $ThrottleLimit"
    Write-Output "========================================`n"
    
    Write-Output "Obtaining Graph API token using Managed Identity..."
    Connect-AzAccount -Identity | Out-Null
    $token = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
    Write-Output "Successfully obtained Graph API token`n"
    
    Write-Output "Connecting to Exchange Online using Managed Identity..."
    Connect-ExchangeOnline -ManagedIdentity -Organization "starkworkspace.onmicrosoft.com" -ShowBanner:$false
    Write-Output "Successfully connected to Exchange Online`n"
    
    $headers = @{
        Authorization = "Bearer $token"
        "Content-Type" = "application/json"
    }
    
    function Invoke-GraphGet {
        param([Parameter(Mandatory)] [string] $Uri)
        try {
            return Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers
        }
        catch {
            $resp = $_.Exception.Response
            if ($null -ne $resp) {
                $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
                $body = $reader.ReadToEnd()
                Write-Error "Graph error calling: $Uri"
                Write-Error "Graph error body: $body"
            }
            throw
        }
    }
    
    Write-Output "Retrieving organization SKUs..."
    $skuUri = "https://graph.microsoft.com/v1.0/subscribedSkus"
    $subscribedSkus = Invoke-GraphGet -Uri $skuUri
    Write-Output "Retrieved $($subscribedSkus.value.Count) SKU(s)`n"
    
    Write-Output "Bulk-fetching all user licenses (this may take a few minutes)..."
    $allUsersLicenses = @{}
    $allUsersUri = "https://graph.microsoft.com/v1.0/users?`$select=userPrincipalName,assignedLicenses&`$top=999"
    $usersFetched = 0
    
    do {
        $result = Invoke-GraphGet -Uri $allUsersUri
        foreach ($user in $result.value) {
            if ($user.userPrincipalName) {
                $allUsersLicenses[$user.userPrincipalName.ToLowerInvariant()] = $user.assignedLicenses
                $usersFetched++
            }
        }
        $allUsersUri = $result.'@odata.nextLink'
        if ($usersFetched % 5000 -eq 0 -and $usersFetched -gt 0) {
            Write-Output "  Fetched $usersFetched users so far..."
        }
    } while ($allUsersUri)
    
    Write-Output "Successfully fetched licenses for $usersFetched users`n"
    
    Write-Output "Building exempt mailbox lookup table (exempt from archiving only)..."
    $exemptFromArchive = @{}
    
    try {
        $exemptMembers = Get-DistributionGroupMember -Identity $ExemptSecurityGroup -ResultSize Unlimited -ErrorAction Stop
        
        if ($exemptMembers) {
            $exemptCount = 0
            foreach ($member in $exemptMembers) {
                if ($member.PrimarySmtpAddress) {
                    $exemptFromArchive[$member.PrimarySmtpAddress.ToString().ToLowerInvariant()] = $true
                    $exemptCount++
                }
            }
            Write-Output "Found $exemptCount mailbox(es) exempt from archiving (retention policy will still be applied)"
            $exemptMembers = $null
        }
        else {
            Write-Output "No exempt members found in the security group"
        }
    }
    catch {
        Write-Warning "Could not retrieve exempt members: $($_.Exception.Message)"
        Write-Output "Continuing without exemptions..."
    }
    
    [System.GC]::Collect()
    
    Write-Output "`nFetching all mailboxes..."
    $allMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum)
    Write-Output "Found $($allMailboxes.Count) mailboxes to process`n"
    
    Write-Output "Starting parallel processing with $ThrottleLimit concurrent threads..."
    Write-Output "========================================`n"
    
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $runspacePool.Open()
    
    $scriptBlock = {
        param(
            $Mailbox,
            $AllUsersLicenses,
            $SubscribedSkus,
            $ExemptFromArchive,
            $RetentionPolicyName,
            $WhatIf
        )
        
        $result = @{
            EmailAddress = $Mailbox.PrimarySmtpAddress.ToString()
            DisplayName = $Mailbox.DisplayName
            Status = "Unknown"
            RetentionStatus = "Unknown"
            ArchiveStatus = "Unknown"
            ErrorMessage = $null
        }
        
        try {
            $emailAddress = $Mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant()
            $userPrincipalName = $Mailbox.UserPrincipalName
            $isSharedMailbox = ($Mailbox.RecipientTypeDetails -eq 'SharedMailbox')
            $isExemptFromArchive = $ExemptFromArchive.ContainsKey($emailAddress)
            
            $validSkus = @('SPE_E5', 'SPE_E3', 'SPE_F5', 'ENTERPRISEPREMIUM', 'ENTERPRISEPACK', 'EXCHANGEENTERPRISE',"SPE_F5_SECCOMP","EXCHANGEARCHIVE_ADDON")
            $hasValidLicense = $false
            $userLicenses = @()
            
            $upnLower = $userPrincipalName.ToLowerInvariant()
            if ($AllUsersLicenses.ContainsKey($upnLower)) {
                $assignedLicenses = $AllUsersLicenses[$upnLower]
                
                foreach ($assigned in $assignedLicenses) {
                    $sku = $SubscribedSkus.value | Where-Object { $_.skuId -eq $assigned.skuId }
                    if ($sku) {
                        $userLicenses += $sku.skuPartNumber
                        if ($validSkus -contains $sku.skuPartNumber) {
                            $hasValidLicense = $true
                        }
                    }
                }
            }
            
            if ($isSharedMailbox) {
                if ($userLicenses.Count -eq 0) {
                    $hasValidLicense = $true
                    $result.Status = "Shared mailbox (unlicensed)"
                }
                elseif ($userLicenses -contains 'EXCHANGEENTERPRISE') {
                    $hasValidLicense = $true
                    $result.Status = "Shared mailbox (Exchange Online Plan 2)"
                }
                elseif (($userLicenses -contains 'EXCHANGEARCHIVE_ADDON') -and ($userLicenses -contains 'EXCHANGESTANDARD')) {
                    $hasValidLicense = $true
                    $result.Status = "Shared mailbox (Exchange Online Plan 1 + Archiving)"
                }
                elseif ($userLicenses | Where-Object { $validSkus -contains $_ }) {
                    $hasValidLicense = $true
                    $matchedSku = ($userLicenses | Where-Object { $validSkus -contains $_ })[0]
                    $result.Status = "Shared mailbox (licensed with $matchedSku)"
                }
                else {
                    $hasValidLicense = $false
                    $result.Status = "SKIPPED (Shared mailbox with invalid license)"
                    $result.RetentionStatus = "Skipped"
                    $result.ArchiveStatus = "Skipped"
                    return $result
                }
            }
            
            if (-not $hasValidLicense) {
                $result.Status = "SKIPPED (No valid license)"
                $result.RetentionStatus = "Skipped"
                $result.ArchiveStatus = "Skipped"
                return $result
            }
            
            $retryCount = 0
            $maxRetries = 3
            $retentionSuccess = $false
            
            while (-not $retentionSuccess -and $retryCount -lt $maxRetries) {
                try {
                    $mbxCheck = Get-Mailbox -Identity $Mailbox.PrimarySmtpAddress -ErrorAction Stop
                    
                    if ($mbxCheck.RetentionPolicy -eq $RetentionPolicyName) {
                        $result.RetentionStatus = "Already assigned"
                        $retentionSuccess = $true
                    }
                    else {
                        if ($WhatIf) {
                            Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -RetentionPolicy $RetentionPolicyName -WhatIf -ErrorAction Stop
                            $result.RetentionStatus = "Would assign (WhatIf)"
                            $retentionSuccess = $true
                        }
                        else {
                            Set-Mailbox -Identity $Mailbox.PrimarySmtpAddress -RetentionPolicy $RetentionPolicyName -ErrorAction Stop
                            $result.RetentionStatus = "Assigned"
                            $retentionSuccess = $true
                        }
                    }
                }
                catch {
                    $retryCount++
                    $errorMessage = $_.Exception.Message
                    
                    if ($errorMessage -match "server side error|try again after some time") {
                        if ($retryCount -lt $maxRetries) {
                            Start-Sleep -Seconds 2
                        }
                        else {
                            $result.RetentionStatus = "FAILED (transient error)"
                            $result.ErrorMessage = $errorMessage
                        }
                    }
                    else {
                        $result.RetentionStatus = "FAILED"
                        $result.ErrorMessage = $errorMessage
                        break
                    }
                }
            }
            
            if ($retentionSuccess) {
                if ($isExemptFromArchive) {
                    $result.ArchiveStatus = "Skipped (exempt)"
                }
                else {
                    try {
                        $mbx = Get-Mailbox -Identity $Mailbox.PrimarySmtpAddress -ErrorAction Stop
                        
                        if ($mbx.ArchiveStatus -eq 'Active') {
                            $result.ArchiveStatus = "Already enabled"
                        }
                        else {
                            if ($WhatIf) {
                                Enable-Mailbox -Identity $Mailbox.PrimarySmtpAddress -Archive -WhatIf -ErrorAction Stop
                                $result.ArchiveStatus = "Would enable (WhatIf)"
                            }
                            else {
                                Enable-Mailbox -Identity $Mailbox.PrimarySmtpAddress -Archive -ErrorAction Stop
                                $result.ArchiveStatus = "Enabled"
                            }
                        }
                    }
                    catch {
                        $result.ArchiveStatus = "FAILED"
                        if (-not $result.ErrorMessage) {
                            $result.ErrorMessage = $_.Exception.Message
                        }
                    }
                }
            }
        }
        catch {
            $result.Status = "FAILED (outer)"
            $result.ErrorMessage = $_.Exception.Message
        }
        
        return $result
    }
    
    $jobs = @()
    $processedCount = 0
    $skippedCount = 0
    $skippedNoLicenseCount = 0
    $assignedCount = 0
    $policyAlreadyAssignedCount = 0
    $archiveEnabledCount = 0
    $archiveAlreadyEnabledCount = 0
    $archiveErrorCount = 0
    $errorCount = 0
    $transientErrorCount = 0
    $lastReportTime = Get-Date
    
    foreach ($mailbox in $allMailboxes) {
        $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($mailbox).AddArgument($allUsersLicenses).AddArgument($subscribedSkus).AddArgument($exemptFromArchive).AddArgument($RetentionPolicyName).AddArgument($WhatIf)
        $powershell.RunspacePool = $runspacePool
        
        $jobs += [PSCustomObject]@{
            Pipe = $powershell
            Result = $powershell.BeginInvoke()
        }
        
        if ($jobs.Count -ge 100 -or $mailbox -eq $allMailboxes[-1]) {
            foreach ($job in $jobs) {
                $result = $job.Pipe.EndInvoke($job.Result)
                $job.Pipe.Dispose()
                
                $processedCount++
                
                if ($result.RetentionStatus -eq "Skipped") {
                    if ($result.Status -like "*No valid license*") {
                        $skippedNoLicenseCount++
                    }
                }
                elseif ($result.RetentionStatus -eq "Already assigned") {
                    $policyAlreadyAssignedCount++
                }
                elseif ($result.RetentionStatus -like "Assigned*" -or $result.RetentionStatus -like "Would assign*") {
                    $assignedCount++
                }
                elseif ($result.RetentionStatus -like "FAILED*") {
                    if ($result.RetentionStatus -like "*transient*") {
                        $transientErrorCount++
                    }
                    else {
                        $errorCount++
                    }
                }
                
                if ($result.ArchiveStatus -eq "Skipped (exempt)") {
                    $skippedCount++
                }
                elseif ($result.ArchiveStatus -eq "Already enabled") {
                    $archiveAlreadyEnabledCount++
                }
                elseif ($result.ArchiveStatus -like "Enabled*" -or $result.ArchiveStatus -like "Would enable*") {
                    $archiveEnabledCount++
                }
                elseif ($result.ArchiveStatus -eq "FAILED") {
                    $archiveErrorCount++
                }
                
                if ($result.ErrorMessage) {
                    Write-Warning "$($result.EmailAddress): $($result.ErrorMessage)"
                }
            }
            
            $jobs = @()
            
            if ($processedCount % 500 -eq 0) {
                $currentTime = Get-Date
                $elapsed = ($currentTime - $lastReportTime).TotalSeconds
                $rate = if ($elapsed -gt 0) { [math]::Round(500 / $elapsed, 1) } else { 0 }
                Write-Output "Progress: $processedCount mailboxes processed | Policy Assigned: $assignedCount | Policy Already Set: $policyAlreadyAssignedCount | Archive Enabled: $archiveEnabledCount | Archive Already Active: $archiveAlreadyEnabledCount | Skipped (Exempt): $skippedCount | Skipped (No License): $skippedNoLicenseCount | Errors: $errorCount | Archive Errors: $archiveErrorCount | Rate: $rate/sec"
                $lastReportTime = $currentTime
                [System.GC]::Collect()
            }
        }
    }
    
    $runspacePool.Close()
    $runspacePool.Dispose()
    
    Write-Output "`n========================================"
    Write-Output "Final Summary"
    Write-Output "========================================"
    Write-Output "Total mailboxes processed: $processedCount"
    Write-Output "Exempt from archiving (retention policy applied): $skippedCount"
    Write-Output "No valid license (skipped): $skippedNoLicenseCount"
    Write-Output "  Valid licenses: SPE_E5, SPE_E3, SPE_F5, ENTERPRISEPREMIUM, ENTERPRISEPACK, EXCHANGEENTERPRISE, SPE_F5_SECCOMP, EXCHANGEARCHIVE_ADDON"
    Write-Output ""
    Write-Output "Retention Policy Enablement Results:"
    Write-Output "  Mailboxes assigned policy: $assignedCount"
    Write-Output "  Mailboxes already assigned policy: $policyAlreadyAssignedCount"
    Write-Output "  Policy assignment errors: $errorCount"
    Write-Output "  Transient errors (after retries): $transientErrorCount"
    Write-Output ""
    Write-Output "Archive Enablement Results:"
    Write-Output "  Archives enabled: $archiveEnabledCount"
    Write-Output "  Archives already active: $archiveAlreadyEnabledCount"
    Write-Output "  Archive enablement errors: $archiveErrorCount"
    Write-Output "WhatIf Mode: $WhatIf"
    Write-Output "========================================`n"
    
    Write-Output "Disconnecting from Exchange Online..."
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Output "Successfully disconnected from Exchange Online"
    
    Write-Output "Disconnecting from Azure Account..."
    Disconnect-AzAccount | Out-Null
    Write-Output "Successfully disconnected from Azure Account"
    
    Write-Output "`nScript completed successfully"
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.ScriptStackTrace)"
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Warning "Could not disconnect from services"
    }
    
    throw
}
