# Mock-tenant harness for the review script's threat-protection section.
$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '..\Invoke-ExchangeOnlineReview.ps1'
$outDir = Join-Path $env:TEMP 'exo-review-test'
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
"Parse errors: $($errors.Count)"

$funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
"Extracted $($funcs.Count) functions"
. ([scriptblock]::Create(($funcs | ForEach-Object { $_.Extent.Text }) -join "`n`n"))

# ---- Script state -------------------------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CollectionLog = [System.Collections.Generic.List[object]]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
$script:Summary = [ordered]@{}
$script:PurviewConnected = $false
$script:PurviewError = $null
$script:SectionTitles = [ordered]@{ ThreatProtection = 'Threat protection'; ClientAccess = 'Client access'; AlertPolicies = 'Alert policies' }
$script:PurviewConnected = $true
$script:IncludePurview = $true
$script:CasMailboxes = $null
$script:MaxRows = 50

# ---- Fakes --------------------------------------------------------------
$script:FakeSpamDomains = 1..30 | ForEach-Object { "domain$_.example.com" }
function Get-ATPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Strict Preset Security Policy'; State = 'Disabled'; Priority = 1; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-EOPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Standard Preset Security Policy'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @('sg@mockcorp.com'); RecipientDomainIs = @(); HostedContentFilterPolicy = 'Standard Preset Security Policy' }) }
function Get-ATPBuiltInProtectionRule { @([PSCustomObject]@{ Name = 'ATP Built-In Protection Rule'; State = 'Enabled'; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @(); SafeLinksPolicy = 'Built-In Protection Policy'; ExceptIfRecipientDomainIs = @('excluded.example.com') }) }
function Get-AntiPhishPolicy { @([PSCustomObject]@{ Name = 'Office365 AntiPhish Default'; Enabled = $true; EnableTargetedUserProtection = $false; TargetedUsersToProtect = @(); EnableTargetedDomainsProtection = $false; TargetedDomainsToProtect = @(); ExcludedSenders = @(); ExcludedDomains = @(); TargetedUserProtectionAction = 'NoAction'; TargetedUserQuarantineTag = 'AdminOnlyAccessPolicy'; EnableOrganizationDomainsProtection = $true; EnableMailboxIntelligence = $true; EnableMailboxIntelligenceProtection = $false; MailboxIntelligenceProtectionAction = 'NoAction'; MailboxIntelligenceQuarantineTag = $null; EnableSpoofIntelligence = $true; AuthenticationFailAction = 'MoveToJmf'; SpoofQuarantineTag = 'DefaultFullAccessPolicy'; HonorDmarcPolicy = $true; DmarcQuarantineAction = 'Quarantine'; DmarcRejectAction = 'Reject'; PhishThresholdLevel = 2; TargetedDomainProtectionAction = 'NoAction'; TargetedDomainQuarantineTag = $null; EnableFirstContactSafetyTips = $false; EnableSimilarUsersSafetyTips = $false; EnableSimilarDomainsSafetyTips = $false; EnableUnusualCharactersSafetyTips = $false; EnableUnauthenticatedSender = $true; EnableViaTag = $true }) }
function Get-AntiPhishRule { @([PSCustomObject]@{ Name = 'Office365 AntiPhish Default'; AntiPhishPolicy = 'Office365 AntiPhish Default'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-HostedContentFilterPolicy {
    $base = @{ SpamAction = 'MoveToJmf'; HighConfidenceSpamAction = 'Quarantine'; PhishSpamAction = 'Quarantine'; HighConfidencePhishAction = 'Quarantine'; BulkSpamAction = 'MoveToJmf'; BulkThreshold = 7; IncreaseScoreWithImageLinks = 'Off'; AllowedSenders = @(); BlockedSenders = @(); BlockedSenderDomains = @('bad.com'); QuarantineRetentionPeriod = 15; SpamQuarantineTag = 'DefaultFullAccessPolicy'; HighConfidenceSpamQuarantineTag = $null; PhishQuarantineTag = 'AdminOnlyAccessPolicy'; HighConfidencePhishQuarantineTag = $null; BulkQuarantineTag = $null;
               EnableLanguageBlockList = $false; LanguageBlockList = @(); EnableRegionBlockList = $false; RegionBlockList = @(); IntraOrgFilterState = 'Default'; InlineSafetyTipsEnabled = $false; PhishZapEnabled = $true; SpamZapEnabled = $true; TestModeAction = 'None';
               IncreaseScoreWithNumericIps = 'Off'; IncreaseScoreWithRedirectToOtherPort = 'Off'; IncreaseScoreWithBizOrInfoUrls = 'Off'; MarkAsSpamEmptyMessages = 'Off'; MarkAsSpamJavaScriptInHtml = 'Off'; MarkAsSpamFramesInHtml = 'Off'; MarkAsSpamObjectTagsInHtml = 'Off'; MarkAsSpamEmbedTagsInHtml = 'Off'; MarkAsSpamFormTagsInHtml = 'Off'; MarkAsSpamWebBugsInHtml = 'Off'; MarkAsSpamSensitiveWordList = 'Off'; MarkAsSpamSpfRecordHardFail = 'Off'; MarkAsSpamFromAddressAuthFail = 'Off'; MarkAsSpamNdrBackscatter = 'Off' }
    @(
        [PSCustomObject]($base + @{ Name = 'Default'; IsDefault = $true; Enabled = $true; AllowedSenderDomains = $script:FakeSpamDomains; MarkAsSpamBulkMail = 'Off' }),
        [PSCustomObject]($base + @{ Name = 'Custom Contoso'; IsDefault = $false; Enabled = $true; AllowedSenderDomains = @(); MarkAsSpamBulkMail = 'On' }),
        [PSCustomObject]($base + @{ Name = 'Disabled Scoped'; IsDefault = $false; Enabled = $true; AllowedSenderDomains = @() }),
        [PSCustomObject]($base + @{ Name = 'Standard Preset Security Policy'; IsDefault = $false; Enabled = $true; RecommendedPolicyType = 'Standard'; AllowedSenderDomains = @() })
    )
}
function Get-HostedContentFilterRule {
    @(
        [PSCustomObject]@{ Name = 'Custom Contoso'; HostedContentFilterPolicy = 'Custom Contoso'; State = 'Enabled'; Priority = 1; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @('contoso.eu'); ExceptIfSentTo = @('user@contoso.eu'); ExceptIfSentToMemberOf = @(); ExceptIfRecipientDomainIs = @() },
        [PSCustomObject]@{ Name = 'Disabled Scoped'; HostedContentFilterPolicy = 'Disabled Scoped'; State = 'Disabled'; Priority = 2; SentTo = @(); SentToMemberOf = @('scoped-group@mockcorp.com'); RecipientDomainIs = @(); ExceptIfSentTo = @(); ExceptIfSentToMemberOf = @(); ExceptIfRecipientDomainIs = @() }
    )
}
function Get-HostedConnectionFilterPolicy { @([PSCustomObject]@{ Name = 'Default'; Enabled = $true; IPAllowList = @(); IPBlockList = @('1.2.3.4'); EnableSafeList = $false }) }
function Get-HostedOutboundSpamFilterPolicy {
    @(
        [PSCustomObject]@{ Name = 'Default'; Enabled = $true; RecipientLimitExternalPerHour = 500; RecipientLimitInternalPerHour = 1000; RecipientLimitPerDay = 1000; AutoForwardingMode = 'Automatic'; NotifyOutboundSpam = $false; NotifyOutboundSpamRecipients = @() },
        [PSCustomObject]@{ Name = 'Outbound custom'; Enabled = $true; RecipientLimitExternalPerHour = 200; RecipientLimitInternalPerHour = 400; RecipientLimitPerDay = 800; ActionWhenThresholdReached = 'BlockUser'; AutoForwardingMode = 'On'; BccSuspiciousOutboundMail = $true; BccSuspiciousOutboundAdditionalRecipients = @('bcc@mockcorp.com'); NotifyOutboundSpam = $true; NotifyOutboundSpamRecipients = @('sec@mockcorp.com') }
    )
}
function Get-HostedOutboundSpamFilterRule {
    @([PSCustomObject]@{ Name = 'Outbound custom'; HostedOutboundSpamFilterPolicy = 'Outbound custom'; State = 'Enabled'; Priority = 1; From = @('ext.sender@mockcorp.com'); FromMemberOf = @(); SenderDomainIs = @(); ExceptIfFrom = @(); ExceptIfFromMemberOf = @(); ExceptIfSenderDomainIs = @() })
}
function Get-MalwareFilterPolicy { @([PSCustomObject]@{ Name = 'Default'; Enabled = $true; EnableFileFilter = $true; FileTypes = @('exe','ps1'); FileTypeAction = 'Block'; QuarantineTag = 'AdminOnlyAccessPolicy'; ZapEnabled = $true; EnableInternalSenderAdminNotifications = $false; InternalSenderAdminAddress = $null; EnableExternalSenderAdminNotifications = $false; ExternalSenderAdminAddress = $null; CustomNotifications = $false; CustomFromName = $null; CustomFromAddress = $null }) }
function Get-MalwareFilterRule { @() }
function Get-SafeAttachmentPolicy { @([PSCustomObject]@{ Name = 'Built-In Protection Policy'; Enable = $true; IsBuiltInProtection = $true; Action = 'Block'; QuarantineTag = 'AdminOnlyAccessPolicy'; Redirect = $false; RedirectAddress = $null; EnableBlockingEncryptedAttachments = $true; ExcludedTypesFromBlockingEncryptedAttachments = @('zip'); QuarantineTagForBlockingEncryptedAttachments = 'DefaultFullAccessWithNotificationPolicy' }) }
function Get-SafeAttachmentRule { @() }
function Get-AtpPolicyForO365 { @([PSCustomObject]@{ Name = 'Default'; EnableATPForSPOTeamsODB = $true; EnableSafeDocs = $false; AllowSafeDocsOpen = $false; EnableSafeLinksForO365Clients = $true; TrackClicks = $true; AllowClickThrough = $false }) }
function Get-SafeLinksPolicy { @([PSCustomObject]@{ Name = 'Built-In Protection Policy'; IsEnabled = $true; IsBuiltInProtection = $true; ScanUrls = $true; EnableSafeLinksForEmail = $true; EnableSafeLinksForTeams = $true; EnableSafeLinksForOffice = $true; EnableForInternalSenders = $false; TrackClicks = $true; AllowClickThrough = $false; DoNotRewriteUrls = @(); DeliverMessageAfterScan = $true; DisableUrlRewrite = $false; EnableOrganizationBranding = $false; CustomNotificationText = $null; UseTranslatedNotificationText = $false }) }
function Get-SafeLinksRule { @() }
function Get-TenantAllowBlockListItems { param($ListType, $ListSubType, $ErrorAction) @() }
function Get-TenantAllowBlockListSpoofItems { @() }
function Get-SecOpsOverridePolicy { @([PSCustomObject]@{ Name = 'SecOps Override Policy'; Identity = 'SecOpsOverridePolicy' }) }
function Get-ExoSecOpsOverrideRule { @([PSCustomObject]@{ Name = 'SecOps override'; Policy = 'SecOpsOverridePolicy'; SentTo = @('secops@contoso.com'); SentToMemberOf = @() }) }
function Get-PhishSimOverridePolicy { @([PSCustomObject]@{ Name = 'PhishSim Override Policy'; Identity = 'PhishSimOverridePolicy' }) }
function Get-ExoPhishSimOverrideRule { @([PSCustomObject]@{ Name = 'PhishSim override'; Policy = 'PhishSimOverridePolicy'; Domains = @('sim.example.com'); SenderIpRanges = @('10.0.0.1') }) }
function Get-QuarantinePolicy { param($QuarantinePolicyType) @() }
function Get-EmailTenantSettings { @() }
function Get-ReportSubmissionPolicy { @() }
function Get-TeamsProtectionPolicy { @([PSCustomObject]@{ Name = 'Teams Protection Policy'; ZapEnabled = $true; HighConfidencePhishQuarantineTag = 'AdminOnlyAccessPolicy'; MalwareQuarantineTag = 'DefaultFullAccessPolicy' }) }
function Get-EXOCASMailbox { param($ResultSize, $Properties)
    @(
        [PSCustomObject]@{ DisplayName = 'U1'; PopEnabled = $true; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $false; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = 'Default' },
        [PSCustomObject]@{ DisplayName = 'U2'; PopEnabled = $false; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $true; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = 'Default' },
        [PSCustomObject]@{ DisplayName = 'U3'; PopEnabled = $false; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $null; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = 'Default' }
    )
}
function Get-OwaMailboxPolicy { @() }
function Get-ActiveSyncOrganizationSettings { @() }
function Get-ActiveSyncDeviceAccessRule { @() }
function Get-MobileDeviceMailboxPolicy { @() }
function Get-AuthenticationPolicy { @() }
function Get-ProtectionAlert {
    @(
        [PSCustomObject]@{ Name = 'DLP alert'; Disabled = $false; IsSystemRule = $true; Category = 'DataLossPrevention'; Severity = 'Medium'; NotifyUser = @('sec@mockcorp.com'); ThreatType = 'None' },
        [PSCustomObject]@{ Name = 'Custom alert'; Disabled = $true; IsSystemRule = $false; Category = 'ThreatManagement'; Severity = 'High'; NotifyUser = @(); ThreatType = 'Malware' }
    )
}

# ---- Run ----------------------------------------------------------------
Get-ThreatProtectionSection
Get-ClientAccessSection
Get-AlertPoliciesSection
$md = $script:Report.ToString()
$mdPath = Join-Path $outDir 'threat-protection.md'
$md | Out-File $mdPath -Encoding utf8
"Markdown written: $mdPath"

# ---- Assertions ----------------------------------------------------------
$pass = 0; $fail = 0
function Check { param([bool]$Cond, [string]$Name)
    if ($Cond) { $script:pass++; "  [PASS] $Name" } else { $script:fail++; "  [FAIL] $Name" }
}

Check ($md -match 'On \(default policy\)') "Default shows 'On (default policy)'"
Check ($md -match 'All recipients') "'All recipients' renders"
Check ($md -match 'Custom Contoso.*Enabled') "custom policy shows 'Enabled'"
Check ($md -match 'Domains: contoso\.eu \(with exclusions\)') "'Domains: contoso.eu (with exclusions)'"
Check ($script:CsvData.AntiSpamInboundPolicies[0].PSObject.Properties.Name -contains 'ExcludedUsers' -or ($script:CsvData.AntiSpamInboundPolicies | Get-Member -Name ExcludedUsers)) "CSV rows have Excluded* columns"
Check ($md -match 'Disabled Scoped.*Disabled') "disabled policy shows 'Disabled'"
Check ($md -match '1 group\(s\)') "scoped '1 group(s)'"
Check ($md -match 'Standard Preset Security Policy.*Enabled') "preset shows 'Enabled'"
Check ($md -match 'IncludedSenders|Included senders') "outbound 'Included senders'"
Check ($md -match 'On \(built-in protection\)') "built-in 'On (built-in protection)'"
Check ($md -match 'All recipients \(with exclusions\)') "'All recipients (with exclusions)'"
$badHeaders = @($md -split "`r?`n" | Where-Object { $_ -match '^\| Name \|' -and $_ -match 'AppliesTo' } | Where-Object {
    $cols = $_ -split '\|' | ForEach-Object { $_.Trim() }
    ($cols -contains 'Enabled') -or ($cols -contains 'IsEnabled') -or ($cols -contains 'SentTo') -or ($cols -contains 'SentToMemberOf') -or ($cols -contains 'RecipientDomainIs')
})
Check ($badHeaders.Count -eq 0) "no Enabled/IsEnabled/SentTo* column headers in policy tables (bad: $($badHeaders.Count))"
Check ($md -notmatch '\|\s*IsEnabled\s*\|') "no IsEnabled cells"
$csvRow = $script:CsvData.AntiSpamInboundPolicies | Where-Object { $_.Name -eq 'Custom Contoso' } | Select-Object -First 1
Check ($csvRow.ExcludedUsers -eq 'user@contoso.eu' -or "$($csvRow.ExcludedUsers)" -eq 'user@contoso.eu') "ExcludedUsers contains user@contoso.eu"
Check ($csvRow.Status -eq 'Enabled' -and $csvRow.AppliesTo -match 'contoso\.eu') "CSV row has Status/AppliesTo"
$slNames = @($script:CsvData.SafeLinksPolicies[0].PSObject.Properties.Name)
Check (($slNames -contains 'ScanUrls') -and -not ($slNames -contains 'ScanUrl')) "SafeLinks CSV has ScanUrls, no ScanUrl"
$saNames = @($script:CsvData.SafeAttachmentPolicies[0].PSObject.Properties.Name)
Check (($saNames -contains 'Enable') -and ($saNames -contains 'QuarantineTagForBlockingEncryptedAttachments')) "SA CSV keeps raw Enable + new props"
$spamNames = @($script:CsvData.AntiSpamInboundPolicies[0].PSObject.Properties.Name)
Check (($spamNames -contains 'MarkAsSpamBulkMail') -and ($spamNames -contains 'TestModeAction')) "spam CSV has all 16 ASF columns + TestModeAction"
Check (@($spamNames | Where-Object { $_ -match 'MarkAsSpam|IncreaseScore' }).Count -eq 16) "16 ASF columns present"
$phishNames = @($script:CsvData.AntiPhishPolicies[0].PSObject.Properties.Name)
Check (($phishNames -contains 'SpoofQuarantineTag') -and ($phishNames -contains 'EnableViaTag')) "anti-phish CSV has new props"

# ---- A2/A3/A4/E checks -----------------------------------------------------
$smtpRow = $script:CsvData.ProtocolUsage | Where-Object { $_.Protocol -eq 'SMTP AUTH' }
Check ($null -ne $smtpRow.FollowsOrgSetting -or ($smtpRow.PSObject.Properties.Name -contains 'FollowsOrgSetting')) "ProtocolUsage has FollowsOrgSetting column"
Check ($smtpRow.FollowsOrgSetting -eq 1 -and $smtpRow.Enabled -eq 1 -and $smtpRow.Disabled -eq 1) "SMTP AUTH: 1 enabled/1 disabled/1 follows org"
$atpNames = @($script:CsvData.AtpPolicyForO365[0].PSObject.Properties.Name)
Check (-not ($atpNames -contains 'TrackClicks') -and -not ($atpNames -contains 'EnableSafeLinksForO365Clients') -and ($atpNames -contains 'EnableATPForSPOTeamsODB')) "ATP O365 trimmed to 3 props"
$tppNames = @($script:CsvData.TeamsProtectionPolicy[0].PSObject.Properties.Name)
Check (($tppNames -contains 'ZapEnabled') -and -not ($tppNames -contains 'ZAPForTeamsEnabled') -and ($tppNames -contains 'MalwareQuarantineTag')) "Teams policy: ZapEnabled + tags, no ZAPForTeamsEnabled"
$alert = $script:CsvData.ProtectionAlerts | Where-Object { $_.Name -eq 'DLP alert' }
Check ($alert.Type -eq 'System' -and $alert.Status -eq 'On' -and $alert.ManagedIn -eq 'Purview DLP policy') "alert Type/Status/ManagedIn (DLP)"
$alert2 = $script:CsvData.ProtectionAlerts | Where-Object { $_.Name -eq 'Custom alert' }
Check ($alert2.Type -eq 'Custom' -and $alert2.Status -eq 'Off' -and $alert2.ManagedIn -eq 'Defender portal (Alert policy)') "alert Type/Status/ManagedIn (custom)"
Check ($md -match 'secops@contoso\.com') "SecOps rule SentTo resolved (no \$_ shadowing)"

# ---- A1 second pass: rule lookup failure ------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
function Get-HostedContentFilterRule { throw 'Simulated rule lookup failure' }
Get-ThreatProtectionSection | Out-Null
$md3 = $script:Report.ToString()
Check ($md3 -match 'Unknown \(rule lookup failed\)') "spam rule fetch failure -> 'Unknown (rule lookup failed)'"
Check (@($script:CollectionLog | Where-Object { $_.Section -eq 'HostedContentFilterRule' }).Count -ge 1) "log has HostedContentFilterRule failure"

# ---- String-typed alert -----------------------------------------------------
$script:CsvData = [ordered]@{}
$script:Report = [System.Text.StringBuilder]::new()
function Get-ProtectionAlert { @([PSCustomObject]@{ Name = 'String alert'; Disabled = 'False'; IsSystemRule = 'True'; Category = 'ThreatManagement'; Severity = 'Low'; NotifyUser = @(); ThreatType = 'Malware' }) }
Get-AlertPoliciesSection | Out-Null
$strAlert = $script:CsvData.ProtectionAlerts | Where-Object { $_.Name -eq 'String alert' }
Check ($strAlert.Status -eq 'On' -and $strAlert.Type -eq 'System') "string-typed alert -> Status On, Type System"

# ---- String-typed mailbox counts --------------------------------------------
$script:CsvData = [ordered]@{}
$script:Report = [System.Text.StringBuilder]::new()
$script:Mailboxes = $null
$script:SectionTitles['MailboxHygiene'] = 'Mailbox hygiene'
function Get-EXOMailbox { param($ResultSize, $Properties, $InactiveMailboxOnly, $SoftDeletedMailbox)
    @(
        [PSCustomObject]@{ UserPrincipalName = 'u1@mockcorp.com'; DisplayName = 'U1'; RecipientTypeDetails = 'UserMailbox'; LitigationHoldEnabled = 'False'; AuditEnabled = 'False'; ArchiveStatus = 'None'; AutoExpandingArchiveEnabled = 'False'; RetentionHoldEnabled = 'False'; ForwardingSmtpAddress = $null; ForwardingAddress = $null; DeliverToMailboxAndForward = 'False' },
        [PSCustomObject]@{ UserPrincipalName = 'u2@mockcorp.com'; DisplayName = 'U2'; RecipientTypeDetails = 'UserMailbox'; LitigationHoldEnabled = 'True'; AuditEnabled = 'True'; ArchiveStatus = 'Active'; AutoExpandingArchiveEnabled = 'False'; RetentionHoldEnabled = 'False'; ForwardingSmtpAddress = $null; ForwardingAddress = $null; DeliverToMailboxAndForward = 'False' }
    )
}
function Get-MailboxAuditBypassAssociation { param($ResultSize) @() }
function Get-User { param($ResultSize, $RecipientTypeDetails, $IsVIP, $Identity) @() }
function Get-InboxRule { param($Mailbox, $ErrorAction) @() }
Get-MailboxHygieneSection | Out-Null
$litRow = $script:CsvData.MailboxHoldCounts | Where-Object { $_.State -eq 'Litigation hold enabled' }
Check ($litRow.Count -eq 1) "string-typed LitigationHoldEnabled -> count 1"
$audRow = $script:CsvData.MailboxAuditCounts | Where-Object { $_.State -eq 'Audit enabled' }
Check ($audRow.Count -eq 1) "string-typed AuditEnabled -> count 1"

# ---- Real-tenant fixes pass -------------------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
function Get-EOPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Standard Preset Security Policy'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @(); HostedContentFilterPolicy = 'Standard Preset Security Policy' }) }
function Get-PhishSimOverridePolicy { @([PSCustomObject]@{ Name = 'PhishSim Override Policy'; Identity = 'PhishSimOverridePolicy'; Guid = 'aaaa0000-1111-2222-3333-444455556666' }) }
function Get-ExoPhishSimOverrideRule { @([PSCustomObject]@{ Name = 'PhishSim override'; Policy = 'aaaa0000-1111-2222-3333-444455556666'; Domains = @('guidsim.example.com'); SenderIpRanges = @('10.0.0.1') }) }
Get-ThreatProtectionSection | Out-Null
$md6 = $script:Report.ToString()
$presetRow = $script:CsvData.AntiSpamInboundPolicies | Where-Object { $_.Name -eq 'Standard Preset Security Policy' } | Select-Object -First 1
Check ($presetRow.AppliesTo -eq 'All recipients') "A1 preset rule without conditions -> 'All recipients'"
Check ($md6 -match 'guidsim\.example\.com') "A2 PhishSim rule matched via GUID shows domains"
$evalInfo = Get-PolicyRuleInfo -Policy ([PSCustomObject]@{ Name = 'Evaluation Policy'; Enabled = $true }) -PresetRules @() -BuiltInRules @()
Check ($evalInfo.Kind -eq 'Evaluation' -and $evalInfo.Status -eq 'Not in use (evaluation policy)' -and $evalInfo.AppliesTo -eq 'Not applicable') "B4 'Evaluation Policy' -> 'Not in use (evaluation policy)'"
Check ((Get-RoleGroupDisplayName 'GlobalReaders_-1791705072') -eq 'GlobalReaders_-1791705072 (Entra role: Global Reader)') "A7 GlobalReaders -> Global Reader"

"=== $pass passed, $fail failed ==="
if ($fail -gt 0) { exit 1 }
