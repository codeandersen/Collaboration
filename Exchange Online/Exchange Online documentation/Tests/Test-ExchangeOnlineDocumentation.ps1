# Mock-tenant harness for Invoke-ExchangeOnlineDocumentation.ps1
# Extracts all functions from the script via AST, fakes the Get-* cmdlets,
# runs every Write-*Section plus CSV/appendix writing, and asserts output rules.

$ErrorActionPreference = 'Stop'
$scriptPath = Join-Path $PSScriptRoot '..\Invoke-ExchangeOnlineDocumentation.ps1'
$outDir = Join-Path $env:TEMP 'exo-doc-test'
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# ---- Parse check --------------------------------------------------------
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
"Parse errors: $($errors.Count)"

# ---- Extract functions --------------------------------------------------
$funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
"Extracted $($funcs.Count) functions"
. ([scriptblock]::Create(($funcs | ForEach-Object { $_.Extent.Text }) -join "`n`n"))

# ---- Script state the functions expect ----------------------------------
$script:ScriptVersion = '1.0-harness'
$script:RunStart = Get-Date
$script:CollectionLog = [System.Collections.Generic.List[object]]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
$script:Report = [System.Text.StringBuilder]::new()
$script:PurviewConnected = $true
$script:PurviewError = $null
$script:TenantId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$script:Mailboxes = $null
$script:CasMailboxes = $null
$script:TenantName = 'Mock Corp'
$script:InitialDomain = 'mockcorp.onmicrosoft.com'
$script:CollectedBy = 'admin@mockcorp.onmicrosoft.com'
$script:ExoModuleVersion = '3.10.1'
$script:AcceptedDomains = $null
$script:DkimConfig = $null
$script:DistGroups = $null
$script:DynGroups = $null
$script:M365Groups = $null
$script:InboundConnectors = $null
$script:OutboundConnectors = $null
$script:TransportRules = $null
$script:OrgConfig = [PSCustomObject]@{ DisplayName = 'Mock Corp'; OrganizationId = '11111111-2222-3333-4444-555555555555'; OAuth2ClientProfileEnabled = $true; AuditDisabled = $false; CustomerLockBoxEnabled = $false; MailTipsAllTipsEnabled = $true; MailTipsExternalRecipientsTipsEnabled = $true; MailTipsLargeAudienceThreshold = 25; EwsEnabled = $null; EwsAllowList = $null; FocusedInboxOn = $true; DefaultAuthenticationPolicy = $null; ActivityBasedAuthenticationTimeoutEnabled = $false; SendFromAliasEnabled = $true; AutoExpandingArchiveEnabled = $false; PublicFoldersEnabled = 'Local'; RootPublicFolderMailbox = 'PFMBX1' }

# Params the functions read
$script:CustomerName = 'MockCorp'
$script:DocumentDate = Get-Date '2026-09-29'
$script:Author = 'Harness'
$script:OutputPath = $outDir
$script:UserPrincipalName = 'admin@mockcorp.onmicrosoft.com'
$script:AppId = $null
$script:CertificateThumbprint = $null
$script:Organization = $null
$script:IncludePurview = $true
$script:MaxListItems = 10
$script:MaxTableRows = 25
$script:CsvDelimiter = $null
$script:DnsServer = $null
$script:ExoLogPath = $null

# ---- DNS stubs (DnsClient module not faked) ------------------------------
function Resolve-DnsSafe { param([string]$Name, [string]$Type)
    if ($Type -eq 'MX' -and $Name -notlike '*.onmicrosoft.com') {
        return @([PSCustomObject]@{ NameExchange = "mockcorp-com.mail.protection.outlook.com"; Section = 'Answer'; Type = 'MX' })
    }
    if ($Type -eq 'CNAME' -and $Name -eq 'selector1._domainkey.mockcorp.com') {
        return @([PSCustomObject]@{ NameHost = 'selector1-mockcorp-com._domainkey.mockcorp.onmicrosoft.com.' })
    }
    if ($Type -eq 'CNAME' -and $Name -eq 'selector2._domainkey.mockcorp.com') {
        return @([PSCustomObject]@{ NameHost = 'selector2-mockcorp-com._domainkey.mockcorp.onmicrosoft.com' })
    }
    return @()
}
function Get-TxtRecords { param([string]$Name)
    if ($Name -like '_dmarc.*') { return @('v=DMARC1; p=quarantine; pct=100; rua=mailto:dmarc@mockcorp.com') }
    if ($Name -like '_mta-sts.*') { return @() }
    if ($Name -like '_smtp._tls.*') { return @() }
    return @('v=spf1 include:spf.protection.outlook.com -all')
}

# ---- Fake cmdlets --------------------------------------------------------
function Get-OrganizationConfig { @($script:OrgConfig) }
function Get-AcceptedDomain {
    @(
        [PSCustomObject]@{ DomainName = 'mockcorp.com'; DomainType = 'Authoritative'; Default = $true; InitialDomain = $false },
        [PSCustomObject]@{ DomainName = 'mockcorp.onmicrosoft.com'; DomainType = 'Authoritative'; Default = $false; InitialDomain = $true }
    )
}
function Get-DkimSigningConfig {
    @(
        [PSCustomObject]@{ Domain = 'mockcorp.com'; Enabled = $true; Selector1CNAME = 'selector1-mockcorp-com._domainkey.mockcorp.onmicrosoft.com'; Selector2CNAME = 'selector2-mockcorp-com._domainkey.mockcorp.onmicrosoft.com'; KeySize = 2048; LastChecked = (Get-Date); RotateOnDate = $null; SelectorBeforeRotateOnDate = $null },
        [PSCustomObject]@{ Domain = 'mockcorp.onmicrosoft.com'; Enabled = $true; Selector1CNAME = 's1'; Selector2CNAME = 's2'; KeySize = 2048; LastChecked = $null; RotateOnDate = $null; SelectorBeforeRotateOnDate = $null }
    )
}
function Get-ArcConfig { @([PSCustomObject]@{ Identity = 'Default'; ArcTrustedSealers = @('sealer1.example.com') }) }
function Get-TransportConfig { throw 'Simulated TransportConfig failure' }
function Get-AdminAuditLogConfig { @([PSCustomObject]@{ UnifiedAuditLogIngestionEnabled = $true; AdminAuditLogEnabled = $true; AdminAuditLogAgeLimit = '90.00:00:00' }) }
function Get-ExternalInOutlook { @([PSCustomObject]@{ Identity = 'bf01e4ea-0499-4b0a-ac1c-8943370b76ca'; Enabled = $true; AllowList = @('partner.com') }) }
function Get-InboundConnector {
    @(
        [PSCustomObject]@{ Name = 'Inbound from partner'; ConnectorType = 'Partner'; Enabled = $true; SenderDomains = @('partner.com'); SenderIPAddresses = @('10.0.0.1'); RequireTls = $true; TlsSenderCertificateName = '*.partner.com'; RestrictDomainsToIPAddresses = $false; RestrictDomainsToCertificate = $true; CloudServicesMailEnabled = $false; TreatMessagesAsInternal = $false; EFSkipLastIP = $false; EFSkipIPs = @(); EFUsers = @(); EFTestMode = $false },
        [PSCustomObject]@{ Name = 'Inbound from on-prem'; ConnectorType = 'OnPremises'; Enabled = $true; SenderDomains = @('*'); SenderIPAddresses = @('192.168.1.1'); RequireTls = $true; TlsSenderCertificateName = 'mail.mockcorp.com'; RestrictDomainsToIPAddresses = $true; RestrictDomainsToCertificate = $false; CloudServicesMailEnabled = $false; TreatMessagesAsInternal = $true; EFSkipLastIP = $true; EFSkipIPs = @('192.168.1.1'); EFUsers = @(); EFTestMode = $false }
    )
}
function Get-OutboundConnector {
    @(
        [PSCustomObject]@{ Name = 'Outbound to partner'; ConnectorType = 'Partner'; Enabled = $true; RecipientDomains = @('partner.com'); SmartHosts = @('mail.partner.com'); TlsSettings = 'DomainValidation'; TlsDomain = 'partner.com'; UseMXRecord = $false; CloudServicesMailEnabled = $false; IsTransportRuleScoped = $false; RouteAllMessagesViaOnPremises = $false },
        [PSCustomObject]@{ Name = 'Outbound via smart host'; ConnectorType = 'OnPremises'; Enabled = $false; RecipientDomains = @('*'); SmartHosts = @('relay.mockcorp.com'); TlsSettings = 'EncryptionOnly'; TlsDomain = $null; UseMXRecord = $false; CloudServicesMailEnabled = $false; IsTransportRuleScoped = $true; RouteAllMessagesViaOnPremises = $true }
    )
}
function Get-TransportRule {
    @(
        [PSCustomObject]@{ Name = 'Append disclaimer'; State = 'Enabled'; Priority = 0; Mode = 'Enforce'; Description = 'Adds a disclaimer to outbound mail.'; Comments = 'Created 2020'; SenderAddressLocation = 'Header'; WhenChanged = (Get-Date '2024-01-01'); FromScope = 'InOrganization'; SentToScope = 'NotInOrganization'; ApplyHtmlDisclaimerLocation = 'Append'; ApplyHtmlDisclaimerText = 'Confidential' },
        [PSCustomObject]@{ Name = 'Block executable attachments'; State = 'Enabled'; Priority = 1; Mode = 'Enforce'; Description = 'Quarantines mail with executable attachments.'; Comments = $null; SenderAddressLocation = 'Header'; WhenChanged = (Get-Date '2024-02-01'); AttachmentHasExecutableContent = $true; Quarantine = $true }
    )
}
function Get-RemoteDomain { @([PSCustomObject]@{ Name = 'Default'; DomainName = '*'; AutoForwardEnabled = $false; AutoReplyEnabled = $true; AllowedOOFType = 'External'; TNEFEnabled = $null; CharacterSet = $null; DeliveryReportEnabled = $true; NDREnabled = $true; MeetingForwardNotificationEnabled = $false; NonMimeCharacterSet = $null }) }
function Get-JournalRule { @([PSCustomObject]@{ Name = 'Journal all'; JournalEmailAddress = 'journal@mockcorp.com'; Scope = 'Global'; Recipient = $null; Enabled = $true }) }
function Get-MailUser { param([switch]$HVEAccount, $ResultSize)
    if ($HVEAccount) { return @() }
    @([PSCustomObject]@{ DisplayName = 'Ext User'; PrimarySmtpAddress = 'ext@mockcorp.com'; RecipientTypeDetails = 'MailUser' })
}
function Get-MailContact { param($ResultSize) @() }
function Get-EXOMailbox { param($ResultSize, $Properties, [switch]$InactiveMailboxOnly, [switch]$SoftDeletedMailbox, [switch]$PublicFolder)
    if ($InactiveMailboxOnly -or $SoftDeletedMailbox) { return @() }
    if ($PublicFolder) { return @([PSCustomObject]@{ DisplayName = 'PF MBX 1' }, [PSCustomObject]@{ DisplayName = 'PF MBX 2' }) }
    $mk = foreach ($i in 1..5) {
        [PSCustomObject]@{ DisplayName = "User $i"; UserPrincipalName = "user$i@mockcorp.com"; PrimarySmtpAddress = "user$i@mockcorp.com"; RecipientTypeDetails = 'UserMailbox'; LitigationHoldEnabled = ($i -eq 1); ArchiveStatus = 'Active'; AutoExpandingArchiveEnabled = $false; RetentionHoldEnabled = $false; AuditEnabled = $true; ForwardingSmtpAddress = $null; ForwardingAddress = $null; DeliverToMailboxAndForward = $false; RoleAssignmentPolicy = 'Default Role Assignment Policy'; RetentionPolicy = 'Default MRM Policy'; WhenMailboxCreated = (Get-Date '2020-01-01'); IsInactiveMailbox = $false; ProhibitSendReceiveQuota = '50 GB' }
    }
    $mk += [PSCustomObject]@{ DisplayName = 'Shared Info'; UserPrincipalName = 'info@mockcorp.com'; PrimarySmtpAddress = 'info@mockcorp.com'; RecipientTypeDetails = 'SharedMailbox'; LitigationHoldEnabled = $false; ArchiveStatus = 'None'; AutoExpandingArchiveEnabled = $false; RetentionHoldEnabled = $false; AuditEnabled = $true; ForwardingSmtpAddress = 'smtp:fwd@partner.com'; ForwardingAddress = $null; DeliverToMailboxAndForward = $true; RoleAssignmentPolicy = 'Default Role Assignment Policy'; RetentionPolicy = $null; WhenMailboxCreated = (Get-Date '2019-01-01'); IsInactiveMailbox = $false; ProhibitSendReceiveQuota = '50 GB' }
    $mk += [PSCustomObject]@{ DisplayName = 'Room A'; UserPrincipalName = 'rooma@mockcorp.com'; PrimarySmtpAddress = 'rooma@mockcorp.com'; RecipientTypeDetails = 'RoomMailbox'; LitigationHoldEnabled = $false; ArchiveStatus = 'None'; AutoExpandingArchiveEnabled = $false; RetentionHoldEnabled = $false; AuditEnabled = $false; ForwardingSmtpAddress = $null; ForwardingAddress = $null; DeliverToMailboxAndForward = $false; RoleAssignmentPolicy = 'Default Role Assignment Policy'; RetentionPolicy = 'Default MRM Policy'; WhenMailboxCreated = (Get-Date '2019-01-01'); IsInactiveMailbox = $false; ProhibitSendReceiveQuota = '50 GB' }
    @($mk)
}
function Get-EXOCASMailbox { param($ResultSize, $Properties)
    @(
        [PSCustomObject]@{ DisplayName = 'User 1'; PopEnabled = $false; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $true; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = 'Default' },
        [PSCustomObject]@{ DisplayName = 'User 2'; PopEnabled = $true; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $true; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = 'Default' },
        [PSCustomObject]@{ DisplayName = 'User 3'; PopEnabled = $false; ImapEnabled = $false; EwsEnabled = $true; ActiveSyncEnabled = $true; MapiEnabled = $true; OWAEnabled = $true; SmtpClientAuthenticationDisabled = $null; OwaMailboxPolicy = 'OwaMailboxPolicy-Default'; ActiveSyncMailboxPolicy = $null }
    )
}
function Get-DistributionGroup { param($ResultSize)
    @(
        [PSCustomObject]@{ Name = 'All Staff'; PrimarySmtpAddress = 'staff@mockcorp.com'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = @('admin'); RequireSenderAuthenticationEnabled = $true },
        [PSCustomObject]@{ Name = 'Sec Team'; PrimarySmtpAddress = 'sec@mockcorp.com'; RecipientTypeDetails = 'MailUniversalSecurityGroup'; ManagedBy = @('admin'); RequireSenderAuthenticationEnabled = $true },
        [PSCustomObject]@{ Name = 'Aarhus Rooms'; PrimarySmtpAddress = 'rooms@mockcorp.com'; RecipientTypeDetails = 'RoomList'; ManagedBy = @(); RequireSenderAuthenticationEnabled = $false },
        [PSCustomObject]@{ Name = 'Orphan Group'; PrimarySmtpAddress = 'orphan@mockcorp.com'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = @(); RequireSenderAuthenticationEnabled = $true }
    )
}
function Get-DynamicDistributionGroup { param($ResultSize)
    @([PSCustomObject]@{ Name = 'Dyn Sales'; PrimarySmtpAddress = 'dynsales@mockcorp.com'; RecipientFilter = "Department -eq 'Sales'"; ManagedBy = @('admin') })
}
function Get-UnifiedGroup { param($ResultSize)
    @(
        [PSCustomObject]@{ DisplayName = 'Project X'; Name = 'Project X'; PrimarySmtpAddress = 'projectx@mockcorp.com'; AccessType = 'Private'; ResourceProvisioningOptions = @('Team'); HiddenFromAddressListsEnabled = $false; RequireSenderAuthenticationEnabled = $false; ManagedBy = @('admin'); GroupExternalMemberCount = 3 },
        [PSCustomObject]@{ DisplayName = 'Open Group'; Name = 'Open Group'; PrimarySmtpAddress = 'open@mockcorp.com'; AccessType = 'Public'; ResourceProvisioningOptions = @(); HiddenFromAddressListsEnabled = $true; RequireSenderAuthenticationEnabled = $true; ManagedBy = @(); GroupExternalMemberCount = 0 }
    )
}
function Get-OwaMailboxPolicy { @([PSCustomObject]@{ Name = 'OwaMailboxPolicy-Default'; IsDefault = $true; ClassicAttachmentsEnabled = $true; ExternalImageProxyEnabled = $false; ThirdPartyFileProvidersEnabled = $false; ConditionalAccessPolicy = 'Off'; AdditionalStorageProvidersAvailable = @('Box') }) }
function Get-MobileDeviceMailboxPolicy { @([PSCustomObject]@{ Name = 'Default'; IsDefault = $true; PasswordEnabled = $true; AlphanumericPasswordRequired = $false; MaxInactivityTimeLock = '00:15:00'; AllowNonProvisionableDevices = $false }) }
function Get-ActiveSyncOrganizationSettings { @([PSCustomObject]@{ DefaultAccessLevel = 'Allow'; UserMailInsert = $null; AdminMailRecipients = @() }) }
function Get-ActiveSyncDeviceAccessRule { @() }
function Get-AuthenticationPolicy { @([PSCustomObject]@{ Name = 'Block legacy'; AllowBasicAuthPop = $false; AllowBasicAuthImap = $false; AllowBasicAuthSmtp = $false; AllowBasicAuthActiveSync = $false; AllowBasicAuthWebService = $false; AllowBasicAuthMapi = $false; AllowBasicAuthOfflineAddressBook = $false; AllowBasicAuthRpc = $false; AllowBasicAuthPowershell = $true; AllowBasicAuthAutodiscover = $false; AllowBasicAuthOutlookService = $false; AllowBasicAuthReportingWebServices = $false; BlockLegacyAuthActiveSync = $true; BlockLegacyAuthImap = $true; BlockLegacyAuthMapi = $true; BlockLegacyAuthOfflineAddressBook = $true; BlockLegacyAuthPop = $true; BlockLegacyAuthRpc = $true; BlockLegacyAuthWebServices = $true }) }
function Get-OrganizationRelationship { @([PSCustomObject]@{ Name = 'Partner share'; Enabled = $true; DomainNames = @('partner.com'); FreeBusyAccessEnabled = $true; FreeBusyAccessLevel = 'AvailabilityOnly' }) }
function Get-SharingPolicy { @([PSCustomObject]@{ Name = 'Default Sharing Policy'; Domains = @('*:CalendarSharingFreeBusySimple','Anonymous:CalendarSharingFreeBusyReviewer'); Default = $true; Enabled = $true }) }
function Get-AvailabilityAddressSpace { @() }
function Get-OnPremisesOrganization { @() }
function Get-IntraOrganizationConnector { @() }
function Get-MigrationEndpoint { @() }
function Get-MigrationBatch { @() }

# EOP / Defender fakes
function Get-EOPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Standard Preset Security Policy'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @('sg@mockcorp.com'); RecipientDomainIs = @(); HostedContentFilterPolicy = 'Standard Preset Security Policy' }) }
function Get-ATPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Strict Preset Security Policy'; State = 'Disabled'; Priority = 1; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-ATPBuiltInProtectionRule { @([PSCustomObject]@{ Name = 'ATP Built-In Protection Rule'; State = 'Enabled'; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @(); SafeLinksPolicy = 'Built-In Protection Policy'; ExceptIfRecipientDomainIs = @('excluded.example.com') }) }
function Get-AntiPhishPolicy { @([PSCustomObject]@{ Name = 'Office365 AntiPhish Default'; Enabled = $true; EnableTargetedUserProtection = $false; TargetedUsersToProtect = @(); EnableTargetedDomainsProtection = $false; TargetedDomainsToProtect = @(); ExcludedSenders = @(); ExcludedDomains = @(); TargetedUserProtectionAction = 'NoAction'; TargetedUserQuarantineTag = 'AdminOnlyAccessPolicy'; EnableOrganizationDomainsProtection = $true; EnableMailboxIntelligence = $true; EnableMailboxIntelligenceProtection = $false; MailboxIntelligenceProtectionAction = 'NoAction'; MailboxIntelligenceQuarantineTag = $null; EnableSpoofIntelligence = $true; AuthenticationFailAction = 'MoveToJmf'; SpoofQuarantineTag = 'DefaultFullAccessPolicy'; HonorDmarcPolicy = $true; DmarcQuarantineAction = 'Quarantine'; DmarcRejectAction = 'Reject'; PhishThresholdLevel = 2; TargetedDomainProtectionAction = 'NoAction'; TargetedDomainQuarantineTag = $null; EnableFirstContactSafetyTips = $false; EnableSimilarUsersSafetyTips = $false; EnableSimilarDomainsSafetyTips = $false; EnableUnusualCharactersSafetyTips = $false; EnableUnauthenticatedSender = $true; EnableViaTag = $true }) }
function Get-AntiPhishRule { @([PSCustomObject]@{ Name = 'Office365 AntiPhish Default'; AntiPhishPolicy = 'Office365 AntiPhish Default'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
$script:FakeSpamDomains = 1..30 | ForEach-Object { "domain$_.example.com" }
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
        [PSCustomObject]@{ Name = 'Outbound custom'; Enabled = $true; RecipientLimitExternalPerHour = 0; RecipientLimitInternalPerHour = 400; RecipientLimitPerDay = 800; ActionWhenThresholdReached = 'BlockUser'; AutoForwardingMode = 'On'; BccSuspiciousOutboundMail = $true; BccSuspiciousOutboundAdditionalRecipients = @('bcc@mockcorp.com'); NotifyOutboundSpam = $true; NotifyOutboundSpamRecipients = @('sec@mockcorp.com') }
    )
}
function Get-HostedOutboundSpamFilterRule {
    @([PSCustomObject]@{ Name = 'Outbound custom'; HostedOutboundSpamFilterPolicy = 'Outbound custom'; State = 'Enabled'; Priority = 1; From = @('ext.sender@mockcorp.com'); FromMemberOf = @(); SenderDomainIs = @(); ExceptIfFrom = @(); ExceptIfFromMemberOf = @(); ExceptIfSenderDomainIs = @() })
}
function Get-MalwareFilterPolicy { @([PSCustomObject]@{ Name = 'Default'; Enabled = $true; EnableFileFilter = $true; FileTypes = @('exe','bat','cmd','vbs','js','ps1','msi','scr','com','pif','reg','wsh'); FileTypeAction = 'Block'; QuarantineTag = 'AdminOnlyAccessPolicy'; ZapEnabled = $true; EnableInternalSenderAdminNotifications = $false; InternalSenderAdminAddress = $null; EnableExternalSenderAdminNotifications = $false; ExternalSenderAdminAddress = $null; CustomNotifications = $false; CustomFromName = $null; CustomFromAddress = $null; Action = 'DeleteAttachmentAndUseDefaultAlert' }) }
function Get-MalwareFilterRule { @([PSCustomObject]@{ Name = 'Default'; MalwareFilterPolicy = 'Default'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-SafeAttachmentPolicy {
    @(
        [PSCustomObject]@{ Name = 'Built-In Protection Policy'; Enable = $true; IsBuiltInProtection = $true; Action = 'Allow'; QuarantineTag = 'AdminOnlyAccessPolicy'; Redirect = $false; RedirectAddress = $null; EnableBlockingEncryptedAttachments = $true; ExcludedTypesFromBlockingEncryptedAttachments = @('zip'); QuarantineTagForBlockingEncryptedAttachments = 'DefaultFullAccessWithNotificationPolicy' },
        [PSCustomObject]@{ Name = 'SA Off'; Enable = $false; Action = 'Allow'; QuarantineTag = $null; Redirect = $false; RedirectAddress = $null; EnableBlockingEncryptedAttachments = $true; ExcludedTypesFromBlockingEncryptedAttachments = @(); QuarantineTagForBlockingEncryptedAttachments = 'DefaultFullAccessWithNotificationPolicy' }
    )
}
function Get-SafeAttachmentRule { @([PSCustomObject]@{ Name = 'Built-In Protection Rule'; SafeAttachmentPolicy = 'Built-In Protection Policy'; State = 'Enabled'; Priority = 0 }) }
function Get-AtpPolicyForO365 { @([PSCustomObject]@{ Name = 'Default'; EnableATPForSPOTeamsODB = $true; EnableSafeDocs = $false; AllowSafeDocsOpen = $false; EnableSafeLinksForO365Clients = $true; TrackClicks = $true; AllowClickThrough = $false }) }
function Get-SafeLinksPolicy { @([PSCustomObject]@{ Name = 'Built-In Protection Policy'; IsEnabled = $true; IsBuiltInProtection = $true; ScanUrls = $true; EnableSafeLinksForEmail = $true; EnableSafeLinksForTeams = $true; EnableSafeLinksForOffice = $true; EnableForInternalSenders = $false; TrackClicks = $true; AllowClickThrough = $false; DoNotRewriteUrls = @('https://skip.example.com'); DeliverMessageAfterScan = $true; DisableUrlRewrite = $false; EnableOrganizationBranding = $false; CustomNotificationText = $null; UseTranslatedNotificationText = $false }) }
function Get-SafeLinksRule { @([PSCustomObject]@{ Name = 'Built-In Protection Rule'; SafeLinksPolicy = 'Built-In Protection Policy'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-QuarantinePolicy { param($QuarantinePolicyType)
    if ($QuarantinePolicyType -eq 'GlobalQuarantinePolicy') {
        return @([PSCustomObject]@{ Name = 'GlobalQuarantinePolicy'; EndUserSpamNotificationFrequency = [timespan]'04:00:00'; EndUserSpamNotificationCustomFromAddress = 'quarantine@messaging.microsoft.com'; OrganizationBrandingEnabled = $true; MultiLanguageSenderName = @('en-US: Mock Corp'); MultiLanguageCustomDisclaimer = @('en-US: Be careful') })
    }
    @([PSCustomObject]@{ Name = 'DefaultFullAccessPolicy'; EndUserQuarantinePermissionsValue = 87; ESNEnabled = $true; QuarantineRetentionDays = 15 })
}
function Get-TenantAllowBlockListItems { param($ListType, $ListSubType, $ErrorAction)
    if ($ListSubType -eq 'AdvancedDelivery') { return @([PSCustomObject]@{ Value = 'https://sim.example.com/phish'; ExpirationDate = (Get-Date '2027-01-01') }) }
    if ($ListType -eq 'Sender') { return @([PSCustomObject]@{ ListType = 'Sender'; Value = 'bad@evil.com'; Action = 'Block'; ExpirationDate = (Get-Date '2027-01-01'); Notes = 'phish' }) }
    return @()
}
function Get-TenantAllowBlockListSpoofItems { @() }
function Get-SecOpsOverridePolicy { @() }
function Get-ExoSecOpsOverrideRule { @() }
function Get-PhishSimOverridePolicy { @() }
function Get-ExoPhishSimOverrideRule { @() }
function Get-ReportSubmissionPolicy { @([PSCustomObject]@{ Name = 'DefaultReportSubmissionPolicy'; EnableReportToMicrosoft = $true; ReportChatMessageEnabled = $false; ReportJunkToCustomizedAddress = $false; ReportNotJunkToCustomizedAddress = $false; ReportPhishToCustomizedAddress = $false; ReportJunkAddresses = @(); ReportPhishAddresses = @() }) }
function Get-EmailTenantSettings { @([PSCustomObject]@{ Identity = 'Default'; EnablePriorityAccountProtection = $false }) }
function Get-TeamsProtectionPolicy { @([PSCustomObject]@{ Name = 'Teams Protection Policy'; ZapEnabled = $true; HighConfidencePhishQuarantineTag = 'AdminOnlyAccessPolicy'; MalwareQuarantineTag = 'DefaultFullAccessPolicy' }) }
function Get-User { param([switch]$IsVIP, $ResultSize)
    if ($IsVIP) { return @([PSCustomObject]@{ DisplayName = 'CEO One'; UserPrincipalName = 'ceo@mockcorp.com' }) }
    @()
}
function Get-MailboxPlan { @([PSCustomObject]@{ DisplayName = 'ExchangeOnline (default)'; IsDefault = $true; ProhibitSendReceiveQuota = '49.5 GB (53,150,220,288 bytes)'; ProhibitSendQuota = '49 GB (52,613,349,376 bytes)'; IssueWarningQuota = '48 GB'; MaxSendSize = '35 MB'; MaxReceiveSize = '36 MB'; RetainDeletedItemsFor = '14.00:00:00'; RetentionPolicy = 'Default MRM Policy'; RoleAssignmentPolicy = 'Default Role Assignment Policy' }) }
function Get-AddressBookPolicy { @([PSCustomObject]@{ Name = 'All Fabrikam ABP'; GlobalAddressList = 'All Fabrikam GAL'; OfflineAddressBook = 'All Fabrikam OAB'; RoomList = 'All Fabrikam Rooms'; AddressLists = @('AL1','AL2') }) }
function Get-AddressList { @([PSCustomObject]@{ Name = 'All Users'; RecipientFilter = "RecipientType -eq 'UserMailbox'" }) }
function Get-GlobalAddressList { @([PSCustomObject]@{ Name = 'Default Global Address List'; RecipientFilter = "Alias -ne `$null" }) }
function Get-OfflineAddressBook { @([PSCustomObject]@{ Name = 'Default Offline Address Book (Ex2019)'; IsDefault = $true; AddressLists = @('\Default Global Address List') }) }
function Get-App { param([switch]$OrganizationApp) @([PSCustomObject]@{ DisplayName = 'Report Message'; Enabled = $true; DefaultStateForUser = 'Enabled'; ProvidedTo = 'Everyone' }) }
function Get-CalendarProcessing { param($Identity) [PSCustomObject]@{ AutomateProcessing = 'AutoAccept'; BookingWindowInDays = 180; MaximumDurationInMinutes = 1440; AllowConflicts = $false; AllowRecurringMeetings = $true; ProcessExternalMeetingMessages = $false; AddOrganizerToSubject = $false; DeleteSubject = $false; AllBookInPolicy = $true } }
function Get-ConnectionInformation { @([PSCustomObject]@{ UserPrincipalName = 'admin@mockcorp.onmicrosoft.com'; TenantID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' }) }

# RBAC fakes
function Get-RoleGroup { @([PSCustomObject]@{ Name = 'Organization Management'; ManagedBy = @(); Roles = @('Mail Recipients','Mail Enabled Public Folders') }, [PSCustomObject]@{ Name = 'Help Desk'; ManagedBy = @('admin'); Roles = @('View-Only Recipients','User Options') }) }
function Get-RoleGroupMember { param($Identity, $ErrorAction) @([PSCustomObject]@{ Name = 'Admin One' }, [PSCustomObject]@{ Name = 'Admin Two' }) }
function Get-RoleAssignmentPolicy { @([PSCustomObject]@{ Name = 'Default Role Assignment Policy'; IsDefault = $true; AssignedRoles = @([PSCustomObject]@{ Name = 'MyBaseOptions' }, [PSCustomObject]@{ Name = 'MyContactInformation' }) }) }
function Get-ManagementRole { @([PSCustomObject]@{ Name = 'Custom Role 1'; Parent = 'Mail Recipients'; RoleType = 'MailRecipients'; IsEndUserRole = $false; IsRootRole = $false }) }
function Get-ManagementRoleAssignment { param($RoleAssigneeType, $ErrorAction) @([PSCustomObject]@{ Role = 'Mail Recipients'; RoleAssigneeName = 'Admin One'; AssignmentMethod = 'Direct'; CustomRecipientWriteScope = $null; RecipientAdministrativeUnitScope = $null; App = $null; RoleAssignee = $null; RoleAssigneeType = 'User'; CustomResourceScope = $null }) }
function Get-ManagementScope { @() }
function Get-ServicePrincipal { @() }

# Compliance fakes
function Get-RetentionPolicy { @([PSCustomObject]@{ Name = 'Default MRM Policy'; RetentionPolicyTagLinks = @([PSCustomObject]@{ Name = 'Default 2 year move to archive' }); IsDefault = $true }) }
function Get-RetentionPolicyTag { @([PSCustomObject]@{ Name = 'Default 2 year move to archive'; Type = 'All'; RetentionAction = 'MoveToArchive'; AgeLimitForRetention = '730.00:00:00'; RetentionEnabled = $true; IsDefaultModeratedRecoveryPolicyTag = $false }, [PSCustomObject]@{ Name = 'Deleted items 30d'; Type = 'DeletedItems'; RetentionAction = 'DeleteAndAllowRecovery'; AgeLimitForRetention = '30.00:00:00'; RetentionEnabled = $true }) }
function Get-IRMConfiguration { @([PSCustomObject]@{ InternalLicensingEnabled = $false; ExternalLicensingEnabled = $false; AzureRMSLicensingEnabled = $false; TransportDecryptionSetting = 'Disabled'; JournalReportDecryptionEnabled = $false; SearchEnabled = $true }) }
function Get-OMEConfiguration { @([PSCustomObject]@{ OTPEnabled = $true; SocialIdSignIn = $true; ExternalMailExpiryInDays = 30 }) }

# Purview fakes
function Get-RetentionCompliancePolicy {
    @(
        [PSCustomObject]@{ Name = 'Exchange retention'; Mode = 'Enforce'; Enabled = $true; ExchangeLocation = @('All') },
        [PSCustomObject]@{ Name = 'Broken policy'; Mode = 'Enforce'; Enabled = $false; ExchangeLocation = @('All') }
    )
}
function Get-RetentionComplianceRule { param($Policy)
    if ($Policy -eq 'Broken policy') { throw 'Simulated retention rule failure' }
    @([PSCustomObject]@{ Name = 'Keep 3y'; PublishComplianceTag = 'Label A'; RetentionDuration = '1095'; ExpirationDateOption = 'CreationAgeInDays'; RetentionComplianceAction = 'Keep' })
}
function Get-DlpCompliancePolicy { @([PSCustomObject]@{ Name = 'PII DLP'; Mode = 'Enable'; Enabled = $true; ExchangeLocation = @('All'); ExchangeLocationException = @() }) }
function Get-DlpComplianceRule { param($Policy, $ErrorAction)
    @(
        [PSCustomObject]@{ Name = 'Rule A'; BlockAccess = $true; NotifyUser = @('SiteAdmin'); GenerateAlert = @('All'); GenerateIncidentReport = @(); Disabled = $false },
        [PSCustomObject]@{ Name = 'Rule B'; BlockAccess = $false; NotifyUser = @(); GenerateAlert = @(); GenerateIncidentReport = @(); Disabled = $true }
    )
}
function Get-ComplianceTag { @([PSCustomObject]@{ Name = 'Keep 10y'; RetentionAction = 'Keep'; RetentionDuration = '3650'; RetentionType = 'CreationAgeInDays'; IsRecordLabel = $true; EventType = $null; Notes = 'Legal hold' }) }
function Get-Label { @([PSCustomObject]@{ DisplayName = 'Confidential'; Name = 'Confidential'; Priority = 0; Tooltip = 'Internal use'; Comment = 'Mark confidential content'; ContentType = @('File','Email'); Disabled = $false }) }
function Get-LabelPolicy { @() }
function Get-ProtectionAlert { @([PSCustomObject]@{ Name = 'Custom alert'; Disabled = $false; IsSystemRule = $false; Category = 'ThreatManagement'; Severity = 'Medium'; NotifyUser = @('sec@mockcorp.com'); ThreatType = 'Malware' }) }

# ---- Run all sections -----------------------------------------------------
"Running sections..."
Write-DocumentInfoSection
Write-TenantOverviewSection
Write-OrgSettingsSection
Write-DomainsSection
Write-MailFlowSection
Write-RecipientsSection
Write-GroupsSection
Write-ClientAccessSection
Write-SharingSection
Write-HybridSection
Write-EmailSecuritySection
Write-PermissionsSection
Write-ComplianceSection
Write-AppendixSection

$mdPath = Join-Path $outDir 'EXO-Documentation_MockCorp_20260929.md'
$script:Report.ToString() | Out-File -FilePath $mdPath -Encoding utf8
"Markdown written: $mdPath ($($script:Report.Length) chars)"

$csvDir = Join-Path $outDir 'csv'
if ($script:CsvData.Count -gt 0) {
    New-Item -ItemType Directory -Force -Path $csvDir | Out-Null
    foreach ($key in $script:CsvData.Keys) {
        @(ConvertTo-CsvRow -Rows $script:CsvData[$key]) | Export-Csv -Path (Join-Path $csvDir "$key.csv") -NoTypeInformation -Encoding UTF8 -UseCulture
    }
}
"CSV files written: $((Get-ChildItem $csvDir -Filter *.csv).Count)"

# ---- Assertions -----------------------------------------------------------
$md = Get-Content $mdPath -Raw -Encoding UTF8
$pass = 0; $fail = 0
function Check { param([bool]$Cond, [string]$Name)
    if ($Cond) { $script:pass++; "  [PASS] $Name" } else { $script:fail++; "  [FAIL] $Name" }
}

# 1. max table columns <= 4 on every '|' line
$badTables = @($script:Report.ToString() -split "`r?`n" | Where-Object { $_ -match '^\s*\|' } | Where-Object {
    ($_ -split '\|' | Where-Object { $_ -notmatch '^\s*$' }).Count -gt 4
})
Check ($badTables.Count -eq 0) "no Markdown table line has more than 4 columns (bad lines: $($badTables.Count))"

# 2. no <br>
Check ($md -notmatch '<br\s*/?>') "no '<br>' in Markdown"

# 3. list cap text for the 30-domain policy
Check ($md -match '\(\+20 more') "list cap '(+20 more' appears"

# 4. 'Not configured' renders
Check ($md -match 'Not configured') "'Not configured' renders"

# 5. 'Not available' renders for throwing cmdlet + log entry
Check ($md -match 'Not available') "'Not available' renders"
Check (@($script:CollectionLog | Where-Object { $_.Section -eq 'TransportConfig' -and $_.Reason -match 'Simulated' }).Count -ge 1) "collection log has TransportConfig failure"

# 6. CSV files match appendix index
$csvFiles = @(Get-ChildItem $csvDir -Filter *.csv | ForEach-Object { $_.Name })
$indexFiles = @($script:CsvIndex.Values | Where-Object { $_.File -ne 'none' } | ForEach-Object { $_.File })
$missingOnDisk = @($indexFiles | Where-Object { $_ -notin $csvFiles })
$missingInIndex = @($csvFiles | Where-Object { $_ -notin $indexFiles })
Check ($missingOnDisk.Count -eq 0 -and $missingInIndex.Count -eq 0) "CSV files match appendix index (onDisk=$($csvFiles.Count) inIndex=$($indexFiles.Count) missingOnDisk=$($missingOnDisk -join ',') missingInIndex=$($missingInIndex -join ','))"
$noneCount = @($script:CsvIndex.Values | Where-Object { $_.File -eq 'none' }).Count
Check ($noneCount -gt 0) "empty datasets appear as 'none' in index ($noneCount entries)"

# ---- Policy status/scope checks (Get-PolicyRuleInfo) ----------------------
Check ($md -match '\| Default \| On \(default policy\) \|') "Default anti-spam shows 'On (default policy)'"
$defAll = ($md -match 'On \(default policy\)[^\r\n]*\| All recipients') -or ($md -match 'Applies to \| All recipients')
Check $defAll "Default applies to 'All recipients'"
Check ($md -match '\| Custom Contoso \| Enabled \|') "custom policy shows 'Enabled'"
Check ($md -match 'Domains: contoso\.eu \(with exclusions\)') "custom policy 'Domains: contoso.eu (with exclusions)'"
Check ($md -match '\| Excluded users \| user@contoso\.eu \|') "card has 'Excluded users | user@contoso.eu'"
Check ($md -match '\| Disabled Scoped \| Disabled \|') "disabled policy shows 'Disabled'"
Check ($md -match '1 group\(s\)') "disabled policy scoped to '1 group(s)'"
Check ($md -match '\| Standard Preset Security Policy \| Enabled \|') "preset policy shows 'Enabled'"
Check ($md -match '\| Included senders \| ext\.sender@mockcorp\.com \|') "outbound shows 'Included senders'"
Check ($md -match '\| Built-In Protection Policy \| On \(built-in protection\) \|') "built-in shows 'On (built-in protection)'"
Check ($md -match 'All recipients \(with exclusions\)') "built-in applies 'All recipients (with exclusions)'"
$secSegment = ($md -split '## Anti-phishing policies', 2)[1] -split '## Tenant Allow/Block List', 2 | Select-Object -First 1
Check (@(($secSegment -split "`r?`n") -match '^\| Enabled \|').Count -eq 0) "no remaining '| Enabled |' rows in policy cards"
$spamCsv = Get-Content (Join-Path $csvDir 'AntiSpamInboundPolicies.csv') -TotalCount 1
Check ($spamCsv -match 'ExcludedUsers' -and $spamCsv -match 'ExcludedDomains') "anti-spam CSV has Excluded* columns"
$slCsv = Get-Content (Join-Path $csvDir 'SafeLinksPolicies.csv') -TotalCount 1
Check ($slCsv -match 'Status' -and $slCsv -match 'AppliesTo' -and $slCsv -match 'ExcludedUsers') "Safe Links CSV has Status/AppliesTo/Excluded* columns"

# ---- Portal-aligned settings checks ---------------------------------------
Check ($md -match '\| Block unscanned attachments \| Yes \|') "SA card 'Block unscanned attachments | Yes'"
Check ($md -match '\| Quarantine policy \(unscanned attachments\) \| DefaultFullAccessWithNotificationPolicy \|') "SA 'Quarantine policy (unscanned attachments)'"
Check ($md -match '\| Safe Attachments unknown malware response \| Monitor \|') "Enable=true+Action=Allow renders 'Monitor'"
Check ($md -match '\| Safe Attachments unknown malware response \| Off \|') "Enable=false renders 'Off'"
Check ($md -match 'Mark bulk email as spam') "ASF row lists non-Off settings by portal name"
Check ($md -match 'None \(all Off\)') "ASF 'None (all Off)' renders"
Check ($md -match '\| Quarantine policy \(spoof\) \| DefaultFullAccessPolicy \|') "anti-phish 'Quarantine policy (spoof)'"
Check ($md -match '\| Apply real-time URL scanning[^|]*\| Yes \|') "Safe Links ScanUrls renders"

# ---- Batch A/B/C checks ----------------------------------------------------
Check ($md -match '\| Tenant ID \| aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee \|') "Tenant ID row shows mocked GUID"
Check ($md -match 'follows org setting 1') "SMTP AUTH row has 'follows org setting' count"
Check ($md -notmatch 'ZAPForTeamsEnabled' -and $md -notmatch 'Track clicks') "no ZAPForTeamsEnabled/TrackClicks"
Check ($md -notmatch '\| Guid \|') "no 'Guid' row (hybrid card)"
Check ($md -match 'Follow user settings') "TNEF 'Follow user settings' renders"
Check ($md -match 'All messages') "journal 'All messages' renders"
Check ($md -match 'All domains: Calendar: free/busy only') "sharing 'All domains' renders"
Check ($md -match 'Anonymous \(published calendars\): Calendar: all information') "sharing 'Anonymous' renders"
Check ($md -match '\| PII DLP \| On \|') "DLP mode 'On' renders"
Check ($md -match 'Rule A: blocks access, notifies user, generates alert') "DLP rule actions render"
Check ($md -match '3650 days \(10 year\(s\)\)') "retention label '3650 days (10 year(s))'"
Check ($md -match '\| When created \|') "retention label 'When created'"
Check ($md -match 'Basic authentication allowed for') "auth policy 'Basic authentication allowed for'"
Check ($md -match '\| Roles \| Mail Recipients, Mail Enabled Public Folders \|' -or $md -match '\| Roles \| .*Mail Recipients') "role group 'Roles' row lists roles"
Check ($md -match '## Mailbox plans \(defaults for new mailboxes\)') "Mailbox plans section exists"
Check ($md -match '## Address lists and address book policies') "Address lists section exists"
Check ($md -match '## Outlook add-ins deployed by the organization') "Outlook add-ins section exists"
Check ($md -match '## Public folders') "Public folders section exists"
Check ($md -match '## Room calendar processing') "Room calendar processing section exists"
Check ($md -match '\| DKIM selector 1 CNAME in DNS \| Published \(matches\) \|') "DKIM 'Published (matches)'"
Check ($md -match '\| Applies to \|') "overview header 'Applies to' humanised"
Check ($md -notmatch '\| AppliesTo \|') "no '| AppliesTo |' header"
Check ($md -match '## Mailbox policy assignment') "renamed 'Mailbox policy assignment' section"
Check ($md -match '## Groups without owners') "Groups without owners section"
Check ($md -match 'Orphan Group') "ownerless group listed"
Check ($md -match '## Global quarantine notification settings' -or $md -match '### Global quarantine notification settings') "global quarantine settings"
Check ($md -match '4 hours') "quarantine frequency '4 hours'"
Check ($md -match '\| Priority accounts \| 1 \|') "priority accounts count"
Check ($md -match '## Microsoft Teams protection') "Teams protection subheading"
$noObjArray = @()
foreach ($f in (Get-ChildItem $csvDir -Filter *.csv)) {
    if ((Get-Content $f.FullName -Raw -ErrorAction SilentlyContinue) -match 'System\.Object\[\]') { $noObjArray += $f.Name }
}
Check ($noObjArray.Count -eq 0) "no 'System.Object[]' in any CSV (bad: $($noObjArray -join ','))"

# ---- A-batch checks --------------------------------------------------------
Check ($md -match '\(no retention policy\)') "retention spread '(no retention policy)'"
Check ($md -match '\| Mail-enabled security groups \| 1 \|') "tenant overview has 'Mail-enabled security groups'"
Check ($md -match '### Default \(\*\)') "remote domain heading 'Default (*)'"
Check ($md -match '\| Keep deleted items for \| 14 days \|') "timespan '14.00:00:00' -> '14 days'"
Check ($md -match '\| Require sign-in after the device has been inactive for \| 15 minutes \|') "timespan '00:15:00' -> '15 minutes'"
Check ($md -match '49\.5 GB \|' -and $md -notmatch '53,150,220,288') "BQS '49.5 GB (... bytes)' shortened"
Check ($md -match '\| Status \| Mode \|') "Purview retention overview has Status"
Check ($md -match 'Unknown \(rule lookup failed\)') "retention 'Unknown (rule lookup failed)'"
Check (@($script:CollectionLog | Where-Object { $_.Section -match 'RetentionComplianceRule' }).Count -ge 1) "log has RetentionComplianceRule failure"
Check ($md -match 'Keep 3y|Publishes retention label') "other retention policy still rendered"
$noOwnerCsv = Get-Content (Join-Path $csvDir 'GroupsNoOwners.csv') -Raw -ErrorAction SilentlyContinue
Check ($noOwnerCsv -notmatch 'RoomList') "no RoomList rows in GroupsNoOwners"
$extCsv = Get-Content (Join-Path $csvDir 'GroupsAcceptExternal.csv') -Raw -ErrorAction SilentlyContinue
Check ($extCsv -notmatch 'RoomList') "no RoomList rows in GroupsAcceptExternal"
Check ($md -match "Move message to the recipients' Junk Email folders") "threat action 'MoveToJmf' humanised"
Check ($md -match 'Default \(entire mailbox\)') "MRM tag type 'All' -> 'Default (entire mailbox)'"
Check ($md -match 'Delete and allow recovery') "MRM tag action humanised"
Check ($md -match '### Built-in protection') "'### Built-in protection' heading"
Check ($md -match '### Simulation URLs to allow') "'### Simulation URLs to allow' heading"
Check ($md -match '### SecOps mailboxes' -and $md -match '### Phishing simulation') "advanced delivery subheadings"
Check ($md -match '### Device access rules') "'### Device access rules' heading"
Check ($md -match '0 \(service default\)') "outbound spam '0 (service default)'"
Check ($md -match '\| Mailbox quota \|') "ProhibitSendReceiveQuota header 'Mailbox quota'"

# ---- A1 second pass: rule lookup failure ------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
function Get-HostedContentFilterRule { throw 'Simulated rule lookup failure' }
Write-EmailSecuritySection
$md2 = $script:Report.ToString()
Check ($md2 -match '\| Custom Contoso \| Unknown \(rule lookup failed\) \|') "anti-spam custom policy 'Unknown (rule lookup failed)'"
Check (@($script:CollectionLog | Where-Object { $_.Section -eq 'HostedContentFilterRule' }).Count -ge 1) "log has HostedContentFilterRule failure"

# ---- String-typed values pass ------------------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
function Get-QuarantinePolicy { param($QuarantinePolicyType)
    if ($QuarantinePolicyType -eq 'GlobalQuarantinePolicy') {
        @([PSCustomObject]@{ Name = 'GlobalQuarantinePolicy'; EndUserSpamNotificationFrequency = '1.00:00:00'; EndUserSpamNotificationCustomFromAddress = $null; OrganizationBrandingEnabled = 'False'; MultiLanguageSenderName = @(); MultiLanguageCustomDisclaimer = @() })
    }
    else {
        @([PSCustomObject]@{ Name = 'DefaultFullAccessPolicy'; ESNEnabled = 'False'; EndUserQuarantinePermissionsValue = '0' })
    }
}
function Get-ProtectionAlert { @([PSCustomObject]@{ Name = 'String alert'; Disabled = 'False'; IsSystemRule = 'False'; Category = 'ThreatManagement'; Severity = 'Low'; NotifyUser = @(); ThreatType = 'Malware' }) }
function Get-EXOMailbox { param($ResultSize, $Properties, [switch]$PublicFolder)
    @(
        [PSCustomObject]@{ DisplayName = 'U1'; PrimarySmtpAddress = 'u1@mockcorp.com'; RecipientTypeDetails = 'UserMailbox'; LitigationHoldEnabled = 'False' },
        [PSCustomObject]@{ DisplayName = 'U2'; PrimarySmtpAddress = 'u2@mockcorp.com'; RecipientTypeDetails = 'UserMailbox'; LitigationHoldEnabled = 'True' }
    )
}
function Get-RemoteDomain {
    @(
        [PSCustomObject]@{ Name = 'Default'; DomainName = '*'; AutoForwardEnabled = 'False'; AutoReplyEnabled = 'True'; AllowedOOFType = 'External'; TNEFEnabled = $null; CharacterSet = $null; DeliveryReportEnabled = 'True'; NDREnabled = 'True'; MeetingForwardNotificationEnabled = 'False'; NonMimeCharacterSet = $null },
        [PSCustomObject]@{ Name = 'Partners'; DomainName = 'fabrikam.com'; AutoForwardEnabled = 'False'; AutoReplyEnabled = 'True'; AllowedOOFType = 'External'; TNEFEnabled = 'False'; CharacterSet = $null; DeliveryReportEnabled = 'True'; NDREnabled = 'True'; MeetingForwardNotificationEnabled = 'False'; NonMimeCharacterSet = $null }
    )
}
Write-EmailSecuritySection
Write-ComplianceSection
Write-RecipientsSection
Write-MailFlowSection
Write-PermissionsSection
$md4 = $script:Report.ToString()
Check ($md4 -match 'Send end-user spam notifications every \| 1 day') "global quarantine '1.00:00:00' -> '1 day'"
Check ($md4 -match '## Tenant Allow/Block List' -and $md4 -match '## Microsoft Teams protection') "email security section survived string-typed frequency"
$strAlert = $script:CsvData.ProtectionAlerts | Where-Object { $_.Name -eq 'String alert' }
Check ($strAlert.Status -eq 'On' -and $strAlert.Type -eq 'Custom') "string-typed alert -> Status On, Type Custom"
Check ($md4 -match '\| Litigation hold enabled \| 1 \|') "litigation hold count 1 with string 'True'"
Check ($md4 -match 'Quarantine notification \| Disabled') "ESNEnabled 'False' -> 'Disabled'"
Check ((Format-Value 'True') -eq 'Yes' -and (Format-Value 'False') -eq 'No') "Format-Value 'True'/'False' -> Yes/No"
Check ($md4 -match '### Partners \(fabrikam\.com\)' -and $md4 -match 'Use rich-text format \(TNEF\) \| Never') "remote domain TNEF 'False' -> 'Never'"
Check ($md4 -match '### Default \(\*\)' -and $md4 -match 'Follow user settings') "remote domain TNEF null -> 'Follow user settings'"
Check ($md4 -match '### Mail Recipients') "direct role assignment heading falls back to Role"
$dra = $script:CsvData.DirectRoleAssignments | Select-Object -First 1
Check ($dra.Scope -eq 'Not configured' -and $null -eq $dra.RecipientWriteScope) "empty recipient scope -> 'Not configured'"

# ---- Real-tenant fixes pass -------------------------------------------------
$script:Report = [System.Text.StringBuilder]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
$script:Mailboxes = $null; $script:CasMailboxes = $null; $script:DistGroups = $null
$script:DynGroups = $null; $script:M365Groups = $null; $script:InboundConnectors = $null
$script:OutboundConnectors = $null; $script:TransportRules = $null
$script:DirObjCache = $null; $script:SpCache = $null
function Get-EOPProtectionPolicyRule { @([PSCustomObject]@{ Name = 'Standard Preset Security Policy'; State = 'Enabled'; Priority = 0; SentTo = @(); SentToMemberOf = @(); RecipientDomainIs = @(); HostedContentFilterPolicy = 'Standard Preset Security Policy' }) }
function Get-HostedContentFilterPolicy { @(
    [PSCustomObject]@{ Name = 'Asf policy'; IsDefault = $false; Enabled = $true; SpamAction = 'MoveToJmf'; IncreaseScoreWithImageLinks = 'On'; MarkAsSpamBulkMail = 'Test' },
    [PSCustomObject]@{ Name = 'Standard Preset Security Policy'; IsDefault = $false; Enabled = $true; RecommendedPolicyType = 'Standard'; SpamAction = 'MoveToJmf' }
) }
function Get-HostedContentFilterRule { @([PSCustomObject]@{ Name = 'Asf policy'; HostedContentFilterPolicy = 'Asf policy'; State = 'Enabled'; Priority = 1; SentTo = @('u@mockcorp.com'); SentToMemberOf = @(); RecipientDomainIs = @() }) }
function Get-PhishSimOverridePolicy { @([PSCustomObject]@{ Name = 'PhishSimOverridePolicy'; Identity = 'PhishSimOverridePolicy'; Guid = 'aaaa0000-1111-2222-3333-444455556666' }) }
function Get-ExoPhishSimOverrideRule { @([PSCustomObject]@{ Name = 'PhishSim rule'; Policy = 'aaaa0000-1111-2222-3333-444455556666'; Domains = @('phishsim.example.com'); SenderIpRanges = @('10.0.0.1') }) }
function Get-LabelPolicy { @([PSCustomObject]@{ Name = 'LP1'; Enabled = $true; ExchangeLocation = @(); ModernGroupLocation = @('sg-mock@mockcorp.com'); ExchangeLocationException = @(); ModernGroupLocationException = @() }) }
function Get-ComplianceTag { @([PSCustomObject]@{ Name = 'Keep 3y'; Guid = 'bbbb0000-1111-2222-3333-444455556666'; RetentionAction = 'Keep'; RetentionDuration = '1095'; RetentionType = 'CreationAgeInDays'; IsRecordLabel = $false }) }
function Get-RetentionComplianceRule { param($Policy)
    if ($Policy -eq 'Broken policy') { throw 'Simulated retention rule failure' }
    @([PSCustomObject]@{ Name = 'Pub rule'; PublishComplianceTag = 'bbbb0000-1111-2222-3333-444455556666'; RetentionDuration = '1095'; ExpirationDateOption = 'CreationAgeInDays'; RetentionComplianceAction = 'Keep' })
}
function Get-Mailbox { param($ResultSize, [switch]$PublicFolder)
    if ($PublicFolder) { @([PSCustomObject]@{ DisplayName = 'PF1' }, [PSCustomObject]@{ DisplayName = 'PF2' }) } else { @() }
}
function Test-CmdletAvailable { param([string]$Name) return ($Name -ne 'Get-AddressList') }
function Get-TransportRule { @([PSCustomObject]@{ Name = 'Long desc'; State = 'Enabled'; Priority = 0; Mode = 'Enforce'; Description = ('x' * 600); Comments = ';'; SenderAddressLocation = 'Header'; WhenChanged = (Get-Date '2024-01-01') }) }
function Get-TransportConfig { @([PSCustomObject]@{ Identity = 'Transport Settings'; JournalingReportNdrTo = '<>'; MaxReceiveSize = '36 MB'; MaxSendSize = '35 MB'; MaxRecipientEnvelopeLimit = 500 }) }
function Get-OutboundConnector { @([PSCustomObject]@{ Name = 'Wildcard out'; ConnectorType = 'Partner'; Enabled = $true; RecipientDomains = @('smtp:*;1'); SmartHosts = @('relay.mockcorp.com'); TlsSettings = 'EncryptionOnly'; TlsDomain = $null; UseMXRecord = $false; CloudServicesMailEnabled = $false; IsTransportRuleScoped = $false; RouteAllMessagesViaOnPremises = $false }) }
function Get-InboundConnector { @([PSCustomObject]@{ Name = 'Wildcard in'; ConnectorType = 'Partner'; Enabled = $true; SenderDomains = @('smtp:*;1'); SenderIPAddresses = @('10.0.0.1'); RequireTls = $false; RestrictDomainsToIPAddresses = $false; RestrictDomainsToCertificate = $false; CloudServicesMailEnabled = $false; TreatMessagesAsInternal = $false; EFSkipLastIP = $false; EFSkipIPs = @(); EFUsers = @(); EFTestMode = $false }) }
function Get-CalendarProcessing { param($Identity) [PSCustomObject]@{ AutomateProcessing = 'AutoAccept'; BookingWindowInDays = 180; MaximumDurationInMinutes = 0; AllowConflicts = $false; AllowRecurringMeetings = $true; ProcessExternalMeetingMessages = $false; AddOrganizerToSubject = $false; DeleteSubject = $false; AllBookInPolicy = $true } }
function Get-RoleGroup { @([PSCustomObject]@{ Name = 'GlobalReaders_-1791705072'; ManagedBy = @(); Roles = @('View-Only Recipients') }) }
function Get-RoleGroupMember { param($Identity, $ErrorAction) @([PSCustomObject]@{ Name = 'dddd0000-1111-2222-3333-444455556666'; RecipientType = 'User' }) }
function Get-User { param([switch]$IsVIP, $ResultSize, $Identity, $ErrorAction)
    if ($Identity) { return [PSCustomObject]@{ DisplayName = "$Identity" } }
    if ($IsVIP) { return @([PSCustomObject]@{ DisplayName = 'CEO One'; UserPrincipalName = 'ceo@mockcorp.com' }) }
    @()
}
function Get-Group { param($Identity, $ErrorAction) $null }
function Get-DistributionGroup { param($ResultSize)
    @([PSCustomObject]@{ Name = 'eeee0000-1111-2222-3333-444455556666'; DisplayName = 'Proc'; RecipientTypeDetails = 'MailUniversalDistributionGroup'; ManagedBy = @(); RequireSenderAuthenticationEnabled = $false; PrimarySmtpAddress = 'proc@mockcorp.com' })
}
function Get-EXOMailbox { param($ResultSize, $Properties, [switch]$PublicFolder)
    @([PSCustomObject]@{ DisplayName = 'Room 1'; PrimarySmtpAddress = 'room1@mockcorp.com'; RecipientTypeDetails = 'RoomMailbox' })
}
Write-OrgSettingsSection
Write-MailFlowSection
Write-RecipientsSection
Write-GroupsSection
Write-ClientAccessSection
Write-EmailSecuritySection
Write-PermissionsSection
Write-ComplianceSection
$md5 = $script:Report.ToString()
$presetSeg = ($md5 -split '### Standard Preset Security Policy', 2)[1]
Check ($presetSeg -match '\| Applies to \| All recipients \|') "A1 preset rule without conditions -> 'All recipients'"
$evalInfo = Get-PolicyRuleInfo -Policy ([PSCustomObject]@{ Name = 'Evaluation Policy'; Enabled = $true }) -PresetRules @() -BuiltInRules @()
Check ($evalInfo.Kind -eq 'Evaluation' -and $evalInfo.Status -eq 'Not in use (evaluation policy)' -and $evalInfo.AppliesTo -eq 'Not applicable' -and $null -eq $evalInfo.Priority) "B4 'Evaluation Policy' -> 'Not in use (evaluation policy)'"
Check ($md5 -match 'phishsim\.example\.com') "A2 PhishSim rule matched via GUID shows domains"
Check ($md5 -match 'Published to users and groups \| sg-mock@mockcorp\.com') "A3 ModernGroupLocation under 'Published to users and groups'"
Check ($md5 -match 'Public folder mailboxes \| 2') "A4 'Public folder mailboxes | 2'"
Check ($md5 -match 'Not available: Get-AddressList') "A5 'Not available: Get-AddressList'"
Check ($md5 -match 'Publishes retention label: Keep 3y') "A6 publish GUID mapped to tag name"
Check ($md5 -match 'GlobalReaders_-1791705072 \(Entra role: Global Reader\)') "A7 GlobalReaders -> Global Reader"
Check (@($script:CsvData.RoleGroups | Where-Object { "$($_.UserMembers)" -match 'Unresolved Entra object dddd0000' }).Count -ge 1) "A8 GUID member -> 'Unresolved Entra object'"
Check ($md5 -match '\| Proc \|') "A9 group shows DisplayName 'Proc'"
Check ($md5 -match 'All domains \(\*\)') "A10 'smtp:*;1' -> 'All domains (*)'"
Check ($md5 -match 'Journaling NDR address \| Not configured') "A11 '<>' -> 'Not configured'"
Check ($md5 -match 'Max duration \(minutes\)' -and $md5 -match '\| No limit \|') "A12 MaximumDurationInMinutes 0 -> 'No limit'"
Check ($md5 -match 'full text in TransportRules\.csv') "B6 long description truncated with CSV pointer"
Check ($md5 -notmatch '\| Comments \| ;') "B6 'Comments | ;' trimmed"
Check ($md5 -match "Move message to the recipients' Junk Email folders") "B1 threat action humanised in new pass"
Check ($md5 -match 'Image links to remote sites') "B2 ASF 'On' -> portal name"
Check ($md5 -match 'Mark bulk email as spam \(test\)') "B2 ASF 'Test' -> 'name (test)'"
Check ($md5 -match '0 \(service default\)') "B7 '0 (service default)' in new pass"

"=== $pass passed, $fail failed ==="
if ($fail -gt 0) { exit 1 }
