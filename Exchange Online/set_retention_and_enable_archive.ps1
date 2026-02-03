#Requires -Modules ExchangeOnlineManagement, Az.Accounts

<#
.SYNOPSIS
    Sets retention policy on all mailboxes except those in the exempt security group.

.DESCRIPTION
    This script connects to Exchange Online using managed identity,
    processes mailboxes using streaming to avoid memory issues,
    excludes members of the exempt security group,
    only enables archiving for users with valid SKUs (SPE_E5, SPE_E3, SPE_F5, ENTERPRISEPREMIUM, ENTERPRISEPACK, EXCHANGEENTERPRISE),
    and applies the specified retention policy in WhatIf mode.
    Optimized for large environments (25,000+ mailboxes).

.NOTES
    Designed for Azure Automation Account
    Requires ExchangeOnlineManagement module
    Runs in WhatIf mode by default
    Uses streaming to avoid memory issues
#>

param(
    [string]$ExemptSecurityGroup = "Exempt_OnlineArchive@starkworkspace.onmicrosoft.com",
    [string]$RetentionPolicyName = "STARK Group Default",
    [switch]$WhatIf = $true
)
write-output "Whatif is set to $WhatIf"
try {
    Write-Output "========================================"
    Write-Output "Retention Policy Assignment Script"
    Write-Output "========================================"
    Write-Output "Exempt Security Group: $ExemptSecurityGroup"
    Write-Output "Retention Policy: $RetentionPolicyName"
    Write-Output "WhatIf Mode: $WhatIf"
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
    
    Write-Output "`nFetching mailboxes that need processing..."
    Write-Output "Note: Filtering to only mailboxes without retention policy or inactive archives for faster processing"
    Write-Output "========================================`n"
    
    $allMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum,Retention,Archive | Where-Object {
        $_.RetentionPolicy -ne $RetentionPolicyName -or $_.ArchiveGuid -eq '00000000-0000-0000-0000-000000000000'
    })
    
    Write-Output "Found $($allMailboxes.Count) mailboxes that need attention (already filtered out fully configured mailboxes)`n"
    Write-Output "Processing mailboxes..."
    Write-Output "========================================`n"
    
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
        $processedCount++
        $emailAddress = $mailbox.PrimarySmtpAddress.ToString().ToLowerInvariant()
        $userPrincipalName = $mailbox.UserPrincipalName
        $isSharedMailbox = ($mailbox.RecipientTypeDetails -eq 'SharedMailbox')
        $isExemptFromArchive = $exemptFromArchive.ContainsKey($emailAddress)
        
        try {
            $validSkus = @('SPE_E5', 'SPE_E3', 'SPE_F5', 'ENTERPRISEPREMIUM', 'ENTERPRISEPACK', 'EXCHANGEENTERPRISE',"SPE_F5_SECCOMP","EXCHANGEARCHIVE_ADDON")
            $hasValidLicense = $false
            $userLicenses = @()
            
            $upnLower = $userPrincipalName.ToLowerInvariant()
            if ($allUsersLicenses.ContainsKey($upnLower)) {
                $assignedLicenses = $allUsersLicenses[$upnLower]
                
                foreach ($assigned in $assignedLicenses) {
                    $sku = $subscribedSkus.value | Where-Object { $_.skuId -eq $assigned.skuId }
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
                    Write-Output "Shared mailbox (unlicensed): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                elseif ($userLicenses -contains 'EXCHANGEENTERPRISE') {
                    $hasValidLicense = $true
                    Write-Output "Shared mailbox (Exchange Online Plan 2): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                elseif (($userLicenses -contains 'EXCHANGEARCHIVE_ADDON') -and ($userLicenses -contains 'EXCHANGESTANDARD')) {
                    $hasValidLicense = $true
                    Write-Output "Shared mailbox (Exchange Online Plan 1 + Archiving): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                elseif ($userLicenses | Where-Object { $validSkus -contains $_ }) {
                    $hasValidLicense = $true
                    $matchedSku = ($userLicenses | Where-Object { $validSkus -contains $_ })[0]
                    Write-Output "Shared mailbox (licensed with $matchedSku): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                else {
                    $hasValidLicense = $false
                    Write-Output "SKIPPED (Shared mailbox with invalid license): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress)) - Licenses: $($userLicenses -join ', ')"
                }
            }
            
            if (-not $hasValidLicense) {
                $skippedNoLicenseCount++
                if (-not $isSharedMailbox) {
                    Write-Output "SKIPPED (No valid license): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                continue
            }
            
            $retryCount = 0
            $maxRetries = 3
            $success = $false
            
            while (-not $success -and $retryCount -lt $maxRetries) {
                try {
                    $mbxCheck = Get-Mailbox -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
                    
                    if ($mbxCheck.RetentionPolicy -eq $RetentionPolicyName) {
                        $policyAlreadyAssignedCount++
                        Write-Output "Retention policy already assigned: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                        $success = $true
                    }
                    else {
                        if ($WhatIf) {
                            Write-output "Whatif: Would set retention policy '$RetentionPolicyName' on $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                            Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -RetentionPolicy $RetentionPolicyName -WhatIf -ErrorAction Stop
                            $assignedCount++
                            $success = $true
                        }
                        else {
                            Write-Output "Set retention policy '$RetentionPolicyName' on $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                            Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -RetentionPolicy $RetentionPolicyName -ErrorAction Stop
                            $assignedCount++
                            $success = $true
                        }
                    }
                }
                catch {
                    $retryCount++
                    $errorMessage = $_.Exception.Message
                    
                    if ($errorMessage -match "server side error|try again after some time") {
                        if ($retryCount -lt $maxRetries) {
                            Write-Warning "Transient error for $($mailbox.PrimarySmtpAddress), retry $retryCount of $maxRetries after 2 seconds..."
                            Start-Sleep -Seconds 2
                        }
                        else {
                            $transientErrorCount++
                            Write-Error "TRANSIENT ERROR (max retries): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress)) - $errorMessage"
                        }
                    }
                    else {
                        $errorCount++
                        Write-Error "FAILED: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress)) - $errorMessage"
                        break
                    }
                }
            }
            
            if ($success) {
                if ($isExemptFromArchive) {
                    $skippedCount++
                    Write-Output "  SKIPPED archive enablement (exempt group member): $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress))"
                }
                else {
                    try {
                        $mbx = Get-Mailbox -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
                        
                        if ($mbx.ArchiveGuid -ne '00000000-0000-0000-0000-000000000000') {
                            $archiveAlreadyEnabledCount++
                            Write-Output "  Archive already enabled for $($mailbox.DisplayName)"
                        }
                        else {
                            if ($WhatIf) {
                                Write-Output "  Whatif: Would enable archive for $($mailbox.DisplayName)"
                                Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -Archive -WhatIf -ErrorAction Stop
                                $archiveEnabledCount++
                            }
                            else {
                                Write-Output "  Enabling archive for $($mailbox.DisplayName)"
                                Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -Archive -ErrorAction Stop | Out-Null
                                $archiveEnabledCount++
                                Write-Output "  Archive enabled successfully for $($mailbox.DisplayName)"
                            }
                        }
                    }
                    catch {
                        $archiveErrorCount++
                        Write-Error "  ARCHIVE FAILED: $($mailbox.DisplayName) ($($mailbox.PrimarySmtpAddress)) - $($_.Exception.Message)"
                    }
                }
            }
        }
        catch {
            $errorCount++
            $errorMsg = if ($_.Exception.Message) { $_.Exception.Message } else { "Unknown error" }
            $mbxDisplay = if ($mailbox.DisplayName) { $mailbox.DisplayName } else { "Unknown" }
            $mbxEmail = if ($mailbox.PrimarySmtpAddress) { $mailbox.PrimarySmtpAddress } else { "Unknown" }
            Write-Error "FAILED (outer): $mbxDisplay ($mbxEmail) - Error: $errorMsg"
        }
        
        if ($processedCount % 500 -eq 0) {
            $currentTime = Get-Date
            $elapsed = ($currentTime - $lastReportTime).TotalSeconds
            $rate = if ($elapsed -gt 0) { [math]::Round(500 / $elapsed, 1) } else { 0 }
            Write-Output "Progress: $processedCount mailboxes processed | Policy Assigned: $assignedCount | Policy Already Set: $policyAlreadyAssignedCount | Archive Enabled: $archiveEnabledCount | Archive Already Active: $archiveAlreadyEnabledCount | Skipped (Exempt): $skippedCount | Skipped (No License): $skippedNoLicenseCount | Errors: $errorCount | Archive Errors: $archiveErrorCount | Rate: $rate/sec"
            $lastReportTime = $currentTime
            [System.GC]::Collect()
        }
    }
    
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
