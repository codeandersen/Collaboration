<#
.SYNOPSIS
    Generates a read-only Markdown review report of an Exchange Online tenant.

.DESCRIPTION
    Connects to Exchange Online and collects Exchange Online, Exchange Online
    Protection (EOP) and Defender for Office 365 configuration without making
    any changes. Optionally collects Microsoft Purview data (Security &
    Compliance PowerShell), Microsoft Graph data (Secure Score, Exchange
    Administrator role holders), per-mailbox statistics and a message trace
    summary. Output is a data-only Markdown report with optional CSV exports.

    The script never calls Set-/New-/Remove-/Enable-/Disable- cmdlets.

.PARAMETER CustomerName
    Name of the customer/tenant being reviewed. Used in the report header and
    file name. Mandatory.

.PARAMETER ReviewedBy
    Name of the person performing the review. Optional.

.PARAMETER ReviewDate
    Date of the review. Defaults to today.

.PARAMETER OutputPath
    Folder for the report and CSV exports. Defaults to .\Reports (relative to
    the current directory). Created if it does not exist.

.PARAMETER UserPrincipalName
    UPN to use for the interactive Exchange Online / IPPS connection.

.PARAMETER AppId
    Application (client) ID for app-only certificate authentication. Requires
    -CertificateThumbprint and -Organization. Used for both EXO and IPPS.

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication.

.PARAMETER Organization
    Organization for app-only authentication (e.g. contoso.onmicrosoft.com).

.PARAMETER IncludePurview
    Connects to Security & Compliance PowerShell (Connect-IPPSSession) and
    adds the Purview sections (alert policies, Purview retention, DLP,
    sensitivity labels, retention labels). Off by default.

.PARAMETER IncludeMailboxStatistics
    Collects per-mailbox sizes, last logon and inbox rules that forward
    externally. Slow on large tenants.

.PARAMETER IncludeMessageTrace
    Adds a 7-day inbound/outbound message summary via Get-MessageTraceV2.

.PARAMETER IncludeGraph
    Collects Microsoft Graph data: Secure Score plus Exchange-related
    controls, and holders of the Entra "Exchange Administrator" role.
    Requires Microsoft.Graph modules.

.PARAMETER ExportCsv
    Exports full detail tables as CSV files next to the report.

.PARAMETER MaxRows
    Maximum number of rows shown per Markdown table. Default 50. Longer
    tables are truncated with a "see CSV" note.

.PARAMETER DnsServer
    Optional DNS server for Resolve-DnsName lookups.

.EXAMPLE
    .\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso"
    EXO-only review. Purview sections are marked as Skipped.

.EXAMPLE
    .\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" -IncludePurview -ExportCsv
    Full review including Purview data and CSV exports.

.EXAMPLE
    .\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" -AppId $id -CertificateThumbprint $thumb -Organization contoso.onmicrosoft.com
    App-only certificate authentication.

.EXAMPLE
    .\Invoke-ExchangeOnlineReview.ps1 -CustomerName "Contoso" -IncludePurview -IncludeGraph -IncludeMailboxStatistics -IncludeMessageTrace -ExportCsv
    Full deep scan with all optional data sources.

.NOTES
    Requires: ExchangeOnlineManagement module (v3)
    Optional:   Microsoft.Graph modules (-IncludeGraph)
    Install:    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    Version:    1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomerName,

    [Parameter(Mandatory = $false)]
    [string]$ReviewedBy,

    [Parameter(Mandatory = $false)]
    [datetime]$ReviewDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\Reports",

    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName,

    [Parameter(Mandatory = $false)]
    [string]$AppId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$Organization,

    [Parameter(Mandatory = $false)]
    [switch]$IncludePurview,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeMailboxStatistics,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeMessageTrace,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeGraph,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false)]
    [int]$MaxRows = 50,

    [Parameter(Mandatory = $false)]
    [string]$DnsServer
)

$script:ScriptVersion = "1.0"
$script:RunStart = Get-Date
$script:CollectionLog = [System.Collections.Generic.List[object]]::new()
$script:Summary = [ordered]@{}
$script:CsvData = [ordered]@{}
$script:Report = [System.Text.StringBuilder]::new()
$script:PurviewConnected = $false
$script:PurviewError = $null
$script:GraphConnected = $false
$script:GraphError = $null
$script:Mailboxes = $null
$script:CasMailboxes = $null
$script:TenantName = $null
$script:InitialDomain = $null
$script:CollectedBy = $UserPrincipalName
$script:SectionTitles = [ordered]@{
    Summary          = 'Summary counts'
    OrgConfig        = 'Organization configuration'
    Recipients       = 'Recipients'
    MailboxHygiene   = 'Mailbox hygiene'
    Groups           = 'Groups'
    Domains          = 'Domains and email authentication'
    MailFlow         = 'Mail flow'
    Hybrid           = 'Hybrid and migration'
    Sharing          = 'Sharing'
    ClientAccess     = 'Client access'
    Permissions      = 'Permissions and RBAC'
    ThreatProtection = 'Threat protection (EOP / Defender for Office 365)'
    AlertPolicies    = 'Alert policies (Purview)'
    Compliance       = 'Compliance and retention (EXO)'
    Graph            = 'Microsoft Graph (opt-in)'
    Appendix         = 'Appendix - collection log'
}

# ============================================================
# Helpers
# ============================================================

function Initialize-ExchangeOnlineModule {
    Write-Host "Checking for ExchangeOnlineManagement module..." -ForegroundColor Cyan
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            Write-Host "Module installed successfully." -ForegroundColor Green
            $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        }
        catch {
            Write-Error "Failed to install ExchangeOnlineManagement module: $_"
            exit 1
        }
    }
    else {
        Write-Host "ExchangeOnlineManagement module $($module.Version) found." -ForegroundColor Green
    }
    $script:ExoModuleVersion = $module.Version.ToString()
}

function Add-Line {
    param([string]$Text = "")
    [void]$script:Report.AppendLine($Text)
}

function Add-LogEntry {
    param([string]$Section, [string]$Reason)
    $script:CollectionLog.Add([PSCustomObject]@{
        Section = $Section
        Reason  = $Reason
        Time    = (Get-Date).ToString("HH:mm:ss")
    })
    Write-Host "  [$Section] $Reason" -ForegroundColor DarkYellow
}

function Test-CmdletAvailable {
    param([string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Assert-Cmdlet {
    param([string]$Name)
    if (-not (Test-CmdletAvailable $Name)) {
        throw "cmdlet $Name not available (module not loaded or workload not licensed)"
    }
}

function ConvertTo-MdAnchor {
    # GitHub slug rules: lowercase, strip anything that is not a letter, digit,
    # space, hyphen or underscore, then replace each space with a hyphen.
    param([string]$Title)
    $slug = $Title.ToLower() -replace '[^a-z0-9 \-_]', ''
    return ($slug -replace ' ', '-')
}

function Convert-ValueToText {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return (@($Value) | ForEach-Object { Convert-ValueToText $_ }) -join "; "
    }
    $text = "$Value"
    $text = $text -replace "\|", "\|"
    $text = $text -replace "`r`n", "<br>" -replace "`n", "<br>" -replace "`r", "<br>"
    return $text
}

function ConvertTo-MdTable {
    param(
        [object[]]$Rows,
        [string[]]$Columns
    )
    if (-not $Rows -or $Rows.Count -eq 0) { return "_None found_" }
    if (-not $Columns -or $Columns.Count -eq 0) {
        $Columns = @($Rows[0].PSObject.Properties.Name)
    }
    $total = $Rows.Count
    $shown = $Rows
    $truncated = $false
    if ($total -gt $MaxRows) {
        $shown = @($Rows | Select-Object -First $MaxRows)
        $truncated = $true
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("| " + ($Columns -join " | ") + " |")
    [void]$sb.AppendLine("|" + (($Columns | ForEach-Object { "---" }) -join "|") + "|")
    foreach ($row in $shown) {
        $cells = foreach ($col in $Columns) { Convert-ValueToText $row.$col }
        [void]$sb.AppendLine("| " + ($cells -join " | ") + " |")
    }
    if ($truncated) {
        $csvNote = if ($ExportCsv) { "see CSV export (-ExportCsv) for full list" } else { "use -ExportCsv for full list" }
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("_Showing $($shown.Count) of $total rows - $csvNote._")
    }
    return $sb.ToString().TrimEnd()
}

function Register-Csv {
    param([string]$Name, [object[]]$Rows)
    if ($Rows -and $Rows.Count -gt 0) {
        $script:CsvData[$Name] = $Rows
    }
}

function Out-Table {
    # Appends a Markdown table for $Rows and registers the full set for CSV export.
    param([string]$CsvName, [object[]]$Rows, [string[]]$Columns)
    Register-Csv -Name $CsvName -Rows $Rows
    Add-Line (ConvertTo-MdTable -Rows $Rows -Columns $Columns)
    Add-Line
}

function Get-PurviewState {
    # $null = collect; otherwise returns the status text to render.
    if (-not $IncludePurview) { return "Skipped" }
    if (-not $script:PurviewConnected) { return "Not available: could not connect to Security & Compliance PowerShell - $script:PurviewError" }
    return $null
}

function Add-PurviewStatusOrThrow {
    # Returns $true when the Purview section should stop and render status text.
    param([string]$Section)
    $state = Get-PurviewState
    if ($null -eq $state) { return $false }
    if ($state -eq "Skipped") {
        Add-Line "_Skipped - Purview data not collected (run with -IncludePurview)_"
    }
    else {
        Add-Line "_$($state)_"
        Add-LogEntry -Section $Section -Reason $state
    }
    return $true
}

function Invoke-Section {
    param(
        [string]$Title,
        [int]$Level = 2,
        [scriptblock]$Body
    )
    Add-Line (('#' * $Level) + " " + $Title)
    Add-Line
    try {
        $result = & $Body
        if ($result) {
            if ($result -is [string]) { Add-Line $result } else { Add-Line ($result | Out-String) }
        }
    }
    catch {
        $reason = $_.Exception.Message
        Add-Line "_Not available: $($reason)_"
        Add-Line
        Add-LogEntry -Section $Title -Reason $reason
    }
    Add-Line
}

function Resolve-DnsSafe {
    param([string]$Name, [string]$Type)
    $params = @{ Name = $Name; Type = $Type; ErrorAction = "Stop" }
    if ($DnsServer) { $params.Server = $DnsServer }
    try { return @(Resolve-DnsName @params) }
    catch { return @() }
}

function Get-TxtRecords {
    param([string]$Name)
    $records = Resolve-DnsSafe -Name $Name -Type TXT
    return @($records | ForEach-Object { ($_.Strings -join "") })
}

function Get-MessageDirection {
    # Classifies a message against the tenant's accepted domains (lowercase list).
    param([string]$Sender, [string]$Recipient, [string[]]$Domains)
    $s = ("$Sender" -split '@')[-1].ToLower()
    $r = ("$Recipient" -split '@')[-1].ToLower()
    $sIn = $Domains -contains $s
    $rIn = $Domains -contains $r
    if ($sIn -and $rIn) { return 'Internal' }
    if ($sIn) { return 'Outbound' }
    if ($rIn) { return 'Inbound' }
    return 'Other'
}

function Get-SharedMailboxes {
    if ($null -eq $script:Mailboxes) {
        $props = @('RecipientTypeDetails', 'UserPrincipalName', 'PrimarySmtpAddress', 'DisplayName',
                   'LitigationHoldEnabled', 'ArchiveStatus', 'AutoExpandingArchiveEnabled', 'RetentionHoldEnabled',
                   'AuditEnabled', 'ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward',
                   'RoleAssignmentPolicy', 'WhenMailboxCreated', 'IsInactiveMailbox', 'ProhibitSendReceiveQuota')
        $script:Mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $props)
    }
    return $script:Mailboxes
}

function Get-SharedCasMailboxes {
    if ($null -eq $script:CasMailboxes) {
        $script:CasMailboxes = @(Get-EXOCASMailbox -ResultSize Unlimited)
    }
    return $script:CasMailboxes
}

# ============================================================
# Report sections
# ============================================================

function Get-OrgConfigSection {
    Invoke-Section -Title $script:SectionTitles.OrgConfig -Body {
        Assert-Cmdlet Get-OrganizationConfig
        $org = if ($script:OrgConfig) { $script:OrgConfig } else { Get-OrganizationConfig }
        Out-Table -CsvName "OrganizationConfig" -Columns @('Setting','Value') -Rows @(
            [PSCustomObject]@{ Setting = 'DisplayName'; Value = $org.DisplayName }
            [PSCustomObject]@{ Setting = 'OAuth2ClientProfileEnabled (modern auth)'; Value = $org.OAuth2ClientProfileEnabled }
            [PSCustomObject]@{ Setting = 'AuditDisabled'; Value = $org.AuditDisabled }
            [PSCustomObject]@{ Setting = 'CustomerLockBoxEnabled'; Value = $org.CustomerLockBoxEnabled }
            [PSCustomObject]@{ Setting = 'MailTipsAllTipsEnabled'; Value = $org.MailTipsAllTipsEnabled }
            [PSCustomObject]@{ Setting = 'MailTipsExternalRecipientsTipsEnabled'; Value = $org.MailTipsExternalRecipientsTipsEnabled }
            [PSCustomObject]@{ Setting = 'MailTipsLargeAudienceThreshold'; Value = $org.MailTipsLargeAudienceThreshold }
            [PSCustomObject]@{ Setting = 'EwsEnabled'; Value = $org.EwsEnabled }
            [PSCustomObject]@{ Setting = 'EwsAllowList'; Value = $org.EwsAllowList }
            [PSCustomObject]@{ Setting = 'FocusedInboxOn'; Value = $org.FocusedInboxOn }
            [PSCustomObject]@{ Setting = 'DefaultAuthenticationPolicy'; Value = $org.DefaultAuthenticationPolicy }
            [PSCustomObject]@{ Setting = 'ActivityBasedAuthenticationTimeoutEnabled'; Value = $org.ActivityBasedAuthenticationTimeoutEnabled }
        )
    }
    Invoke-Section -Title "External sender tagging" -Level 3 -Body {
        Assert-Cmdlet Get-ExternalInOutlook
        $ext = Get-ExternalInOutlook
        $rows = @($ext | ForEach-Object {
            [PSCustomObject]@{
                Identity          = $_.Identity
                Enabled           = $_.Enabled
                AllowList         = $_.AllowList
            }
        })
        Out-Table -CsvName "ExternalInOutlook" -Rows $rows -Columns @('Identity','Enabled','AllowList')
    }
    Invoke-Section -Title "Transport configuration" -Level 3 -Body {
        Assert-Cmdlet Get-TransportConfig
        $tc = Get-TransportConfig
        Out-Table -CsvName "TransportConfig" -Columns @('Setting','Value') -Rows @(
            [PSCustomObject]@{ Setting = 'SmtpClientAuthenticationDisabled'; Value = $tc.SmtpClientAuthenticationDisabled }
            [PSCustomObject]@{ Setting = 'AllowLegacyTLSClients'; Value = $tc.AllowLegacyTLSClients }
            [PSCustomObject]@{ Setting = 'MaxSendSize'; Value = $tc.MaxSendSize }
            [PSCustomObject]@{ Setting = 'MaxReceiveSize'; Value = $tc.MaxReceiveSize }
            [PSCustomObject]@{ Setting = 'ExternalPostmasterAddress'; Value = $tc.ExternalPostmasterAddress }
            [PSCustomObject]@{ Setting = 'JournalingReportNdrTo'; Value = $tc.JournalingReportNdrTo }
            [PSCustomObject]@{ Setting = 'MaxRecipientEnvelopeLimit'; Value = $tc.MaxRecipientEnvelopeLimit }
        )
    }
    Invoke-Section -Title "Unified Audit Log" -Level 3 -Body {
        Assert-Cmdlet Get-AdminAuditLogConfig
        $aal = Get-AdminAuditLogConfig
        Out-Table -CsvName "AdminAuditLogConfig" -Columns @('Setting','Value') -Rows @(
            [PSCustomObject]@{ Setting = 'UnifiedAuditLogIngestionEnabled'; Value = $aal.UnifiedAuditLogIngestionEnabled }
            [PSCustomObject]@{ Setting = 'AdminAuditLogEnabled'; Value = $aal.AdminAuditLogEnabled }
            [PSCustomObject]@{ Setting = 'AdminAuditLogAgeLimit'; Value = $aal.AdminAuditLogAgeLimit }
        )
    }
}

function Get-RecipientsSection {
    Invoke-Section -Title $script:SectionTitles.Recipients -Body {
        Assert-Cmdlet Get-EXOMailbox
        $mbx = Get-SharedMailboxes
        $rows = @($mbx | Group-Object RecipientTypeDetails | Sort-Object Name | ForEach-Object {
            [PSCustomObject]@{ RecipientType = $_.Name; Count = $_.Count }
        })
        $script:Summary['Mailboxes (total)'] = $mbx.Count
        Out-Table -CsvName "MailboxCountsByType" -Rows $rows -Columns @('RecipientType','Count')

        Add-Line "### Mail users and contacts"
        Add-Line
        $mailUsers = @(Get-MailUser -ResultSize Unlimited)
        $guests = @($mailUsers | Where-Object { $_.RecipientTypeDetails -eq 'GuestMailUser' })
        $contacts = @(Get-MailContact -ResultSize Unlimited)
        $pfMailboxes = @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'PublicFolderMailbox' })
        Out-Table -CsvName "RecipientCounts" -Columns @('Type','Count') -Rows @(
            [PSCustomObject]@{ Type = 'Mail users'; Count = @($mailUsers | Where-Object { $_.RecipientTypeDetails -eq 'MailUser' }).Count }
            [PSCustomObject]@{ Type = 'Guest mail users'; Count = $guests.Count }
            [PSCustomObject]@{ Type = 'Mail contacts'; Count = $contacts.Count }
            [PSCustomObject]@{ Type = 'Mail-enabled public folder mailboxes'; Count = $pfMailboxes.Count }
        )
        $script:Summary['Mail users'] = @($mailUsers | Where-Object { $_.RecipientTypeDetails -eq 'MailUser' }).Count
        $script:Summary['Guest mail users'] = $guests.Count
        $script:Summary['Mail contacts'] = $contacts.Count

        Add-Line "### Inactive and soft-deleted mailboxes"
        Add-Line
        $inactive = @()
        $softDeleted = @()
        try { $inactive = @(Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Inactive mailboxes' -Reason $_.Exception.Message }
        try { $softDeleted = @(Get-EXOMailbox -SoftDeletedMailbox -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Soft-deleted mailboxes' -Reason $_.Exception.Message }
        Out-Table -CsvName "InactiveMailboxes" -Columns @('Type','Count') -Rows @(
            [PSCustomObject]@{ Type = 'Inactive mailboxes'; Count = $inactive.Count }
            [PSCustomObject]@{ Type = 'Soft-deleted mailboxes'; Count = $softDeleted.Count }
        )
        $script:Summary['Inactive mailboxes'] = $inactive.Count
        $script:Summary['Soft-deleted mailboxes'] = $softDeleted.Count
    }
}

function Get-MailboxHygieneSection {
    Invoke-Section -Title $script:SectionTitles.MailboxHygiene -Body {
        $mbx = Get-SharedMailboxes

        Add-Line "### Holds and archiving"
        Add-Line
        Out-Table -CsvName "MailboxHoldCounts" -Columns @('State','Count') -Rows @(
            [PSCustomObject]@{ State = 'Litigation hold enabled'; Count = @($mbx | Where-Object { $_.LitigationHoldEnabled }).Count }
            [PSCustomObject]@{ State = 'Archive enabled'; Count = @($mbx | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count }
            [PSCustomObject]@{ State = 'Auto-expanding archive'; Count = @($mbx | Where-Object { $_.AutoExpandingArchiveEnabled }).Count }
            [PSCustomObject]@{ State = 'Retention hold enabled'; Count = @($mbx | Where-Object { $_.RetentionHoldEnabled }).Count }
        )

        Add-Line "### Mailbox auditing"
        Add-Line
        Out-Table -CsvName "MailboxAuditCounts" -Columns @('State','Count') -Rows @(
            [PSCustomObject]@{ State = 'Audit enabled'; Count = @($mbx | Where-Object { $_.AuditEnabled }).Count }
            [PSCustomObject]@{ State = 'Audit disabled'; Count = @($mbx | Where-Object { -not $_.AuditEnabled }).Count }
        )
        $bypass = @()
        try {
            Assert-Cmdlet Get-MailboxAuditBypassAssociation
            $bypass = @(Get-MailboxAuditBypassAssociation -ResultSize Unlimited | Where-Object { $_.AuditBypassEnabled })
        }
        catch { Add-LogEntry -Section 'Mailbox audit bypass' -Reason $_.Exception.Message }
        $bypassRows = @($bypass | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Name; AuditBypassEnabled = $_.AuditBypassEnabled }
        })
        Add-Line "Accounts with audit bypass enabled:"
        Add-Line
        Out-Table -CsvName "AuditBypass" -Rows $bypassRows -Columns @('Identity','AuditBypassEnabled')

        Add-Line "### Mailbox forwarding"
        Add-Line
        $fwd = @($mbx | Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress } | ForEach-Object {
            [PSCustomObject]@{
                UserPrincipalName         = $_.UserPrincipalName
                ForwardingSmtpAddress     = $_.ForwardingSmtpAddress
                ForwardingAddress         = $_.ForwardingAddress
                DeliverToMailboxAndForward = $_.DeliverToMailboxAndForward
            }
        })
        Out-Table -CsvName "MailboxForwarding" -Rows $fwd -Columns @('UserPrincipalName','ForwardingSmtpAddress','ForwardingAddress','DeliverToMailboxAndForward')
        $script:Summary['Mailboxes with forwarding'] = $fwd.Count

        Add-Line "### Shared/room mailboxes with sign-in enabled"
        Add-Line
        $users = @()
        try { $users = @(Get-User -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,RoomMailbox,EquipmentMailbox) }
        catch { Add-LogEntry -Section 'Shared/room sign-in state' -Reason $_.Exception.Message }
        $enabledRows = @($users | Where-Object { -not $_.AccountDisabled } | ForEach-Object {
            [PSCustomObject]@{ UserPrincipalName = $_.UserPrincipalName; RecipientTypeDetails = $_.RecipientTypeDetails; AccountDisabled = $_.AccountDisabled }
        })
        Out-Table -CsvName "SharedMailboxesSignInEnabled" -Rows $enabledRows -Columns @('UserPrincipalName','RecipientTypeDetails','AccountDisabled')
        $script:Summary['Shared/room mailboxes with sign-in enabled'] = $enabledRows.Count

        Add-Line "### Role assignment policy spread"
        Add-Line
        $rapRows = @($mbx | Group-Object RoleAssignmentPolicy | Sort-Object Count -Descending | ForEach-Object {
            [PSCustomObject]@{ RoleAssignmentPolicy = $_.Name; MailboxCount = $_.Count }
        })
        Out-Table -CsvName "RoleAssignmentPolicySpread" -Rows $rapRows -Columns @('RoleAssignmentPolicy','MailboxCount')
    }

    Invoke-Section -Title "Mailbox statistics (-IncludeMailboxStatistics)" -Level 3 -Body {
        if (-not $IncludeMailboxStatistics) {
            Add-Line "_Skipped - mailbox statistics not collected (run with -IncludeMailboxStatistics)_"
            return
        }
        Assert-Cmdlet Get-EXOMailboxStatistics
        $mbx = Get-SharedMailboxes
        $stats = @()
        foreach ($m in $mbx) {
            try {
                $s = Get-EXOMailboxStatistics -Identity $m.UserPrincipalName -Properties LastLogonTime
                $sizeBytes = 0
                if ("$($s.TotalItemSize)" -match '\(([\d,\s]+)\s*bytes\)') { $sizeBytes = [int64]($Matches[1] -replace '[^\d]', '') }
                $stats += [PSCustomObject]@{
                    UserPrincipalName = $m.UserPrincipalName
                    RecipientTypeDetails = $m.RecipientTypeDetails
                    TotalItemSize = $s.TotalItemSize
                    SizeBytes = $sizeBytes
                    ItemCount = $s.ItemCount
                    LastLogonTime = $s.LastLogonTime
                }
            } catch { }
        }
        Register-Csv -Name 'MailboxStatistics' -Rows $stats
        $script:Summary['Mailbox statistics collected'] = $stats.Count

        Add-Line "#### Top mailboxes by size"
        Add-Line
        $top = @($stats | Sort-Object SizeBytes -Descending | Select-Object -First $MaxRows)
        Out-Table -CsvName "TopMailboxesBySize" -Rows $top -Columns @('UserPrincipalName','RecipientTypeDetails','TotalItemSize','ItemCount','LastLogonTime')

        Add-Line "#### Mailboxes not logged on for more than 90 days"
        Add-Line
        $cutoff = (Get-Date).AddDays(-90)
        $stale = @($stats | Where-Object { $_.LastLogonTime -and [datetime]$_.LastLogonTime -lt $cutoff })
        Out-Table -CsvName "StaleMailboxes" -Rows $stale -Columns @('UserPrincipalName','RecipientTypeDetails','LastLogonTime')

        Add-Line "#### Inbox rules forwarding or redirecting externally"
        Add-Line
        $rules = @()
        foreach ($m in $mbx) {
            try {
                $r = Get-InboxRule -Mailbox $m.UserPrincipalName -ErrorAction Stop |
                    Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }
                foreach ($rule in $r) {
                    $targets = @(@($rule.ForwardTo) + @($rule.ForwardAsAttachmentTo) + @($rule.RedirectTo)) -join '; '
                    if ($targets -match '@') {
                        $rules += [PSCustomObject]@{
                            Mailbox = $m.UserPrincipalName
                            RuleName = $rule.Name
                            ForwardTo = $rule.ForwardTo
                            ForwardAsAttachmentTo = $rule.ForwardAsAttachmentTo
                            RedirectTo = $rule.RedirectTo
                            Enabled = $rule.Enabled
                        }
                    }
                }
            } catch { }
        }
        Out-Table -CsvName "ExternalForwardingInboxRules" -Rows $rules -Columns @('Mailbox','RuleName','ForwardTo','ForwardAsAttachmentTo','RedirectTo','Enabled')
        $script:Summary['Inbox rules forwarding externally'] = $rules.Count
    }
}

function Get-GroupsSection {
    Invoke-Section -Title $script:SectionTitles.Groups -Body {
        Assert-Cmdlet Get-DistributionGroup
        $dg = @(Get-DistributionGroup -ResultSize Unlimited)
        $dyn = @()
        try { $dyn = @(Get-DynamicDistributionGroup -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Dynamic distribution groups' -Reason $_.Exception.Message }
        $m365 = @()
        try { $m365 = @(Get-UnifiedGroup -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Microsoft 365 groups' -Reason $_.Exception.Message }

        Add-Line "### Summary"
        Add-Line
        Out-Table -CsvName "GroupCounts" -Columns @('Type','Count') -Rows @(
            [PSCustomObject]@{ Type = 'Distribution groups'; Count = @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup' }).Count }
            [PSCustomObject]@{ Type = 'Mail-enabled security groups'; Count = @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup' }).Count }
            [PSCustomObject]@{ Type = 'Dynamic distribution groups'; Count = $dyn.Count }
            [PSCustomObject]@{ Type = 'Microsoft 365 groups'; Count = $m365.Count }
            [PSCustomObject]@{ Type = 'Room lists'; Count = @($dg | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' }).Count }
        )
        $script:Summary['Distribution groups'] = $dg.Count
        $script:Summary['Dynamic distribution groups'] = $dyn.Count
        $script:Summary['Microsoft 365 groups'] = $m365.Count

        Add-Line "### Dynamic distribution groups"
        Add-Line
        $dynRows = @($dyn | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; PrimarySmtpAddress = $_.PrimarySmtpAddress; RecipientFilter = $_.RecipientFilter; ManagedBy = $_.ManagedBy }
        })
        Out-Table -CsvName "DynamicDistributionGroups" -Rows $dynRows -Columns @('Name','PrimarySmtpAddress','RecipientFilter','ManagedBy')

        Add-Line "### Microsoft 365 groups"
        Add-Line
        $m365Rows = @($m365 | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                PrimarySmtpAddress = $_.PrimarySmtpAddress
                AccessType = $_.AccessType
                HiddenFromExchangeClientsEnabled = $_.HiddenFromExchangeClientsEnabled
                HiddenFromAddressListsEnabled = $_.HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $_.RequireSenderAuthenticationEnabled
                AllowExternalSenders = (-not $_.RequireSenderAuthenticationEnabled)
                ManagedBy = $_.ManagedBy
            }
        })
        Out-Table -CsvName "M365Groups" -Rows $m365Rows -Columns @('DisplayName','PrimarySmtpAddress','AccessType','AllowExternalSenders','RequireSenderAuthenticationEnabled','HiddenFromExchangeClientsEnabled','HiddenFromAddressListsEnabled','ManagedBy')

        Add-Line "### Groups with no owners (ManagedBy)"
        Add-Line
        $noOwner = @($dg | Where-Object { -not $_.ManagedBy -or @($_.ManagedBy).Count -eq 0 } | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientTypeDetails = $_.RecipientTypeDetails; PrimarySmtpAddress = $_.PrimarySmtpAddress }
        })
        $noOwner += @($m365 | Where-Object { -not $_.ManagedBy -or @($_.ManagedBy).Count -eq 0 } | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientTypeDetails = 'GroupMailbox'; PrimarySmtpAddress = $_.PrimarySmtpAddress }
        })
        Out-Table -CsvName "GroupsNoOwners" -Rows $noOwner -Columns @('Name','RecipientTypeDetails','PrimarySmtpAddress')
        $script:Summary['Groups with no owners'] = $noOwner.Count

        Add-Line "### Groups accepting external senders"
        Add-Line
        $extRows = @($dg | Where-Object { -not $_.RequireSenderAuthenticationEnabled } | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientTypeDetails = $_.RecipientTypeDetails; PrimarySmtpAddress = $_.PrimarySmtpAddress; RequireSenderAuthenticationEnabled = $_.RequireSenderAuthenticationEnabled }
        })
        Out-Table -CsvName "GroupsAcceptExternal" -Rows $extRows -Columns @('Name','RecipientTypeDetails','PrimarySmtpAddress','RequireSenderAuthenticationEnabled')
    }
}

function Get-DomainsSection {
    Invoke-Section -Title $script:SectionTitles.Domains -Body {
        Assert-Cmdlet Get-AcceptedDomain
        $domains = if ($script:AcceptedDomains) { $script:AcceptedDomains } else { @(Get-AcceptedDomain) }
        $domRows = @($domains | ForEach-Object {
            [PSCustomObject]@{
                DomainName = $_.DomainName
                DomainType = $_.DomainType
                Default    = $_.Default
                Initial    = $_.InitialDomain
            }
        })
        Out-Table -CsvName "AcceptedDomains" -Rows $domRows -Columns @('DomainName','DomainType','Default','Initial')
        $script:Summary['Accepted domains'] = $domains.Count

        $dkim = @()
        try { $dkim = @(Get-DkimSigningConfig) } catch { Add-LogEntry -Section 'DKIM config' -Reason $_.Exception.Message }

        Add-Line "### DNS records per domain"
        Add-Line
        $dnsRows = @()
        foreach ($d in $domains) {
            $name = $d.DomainName
            if ($name -like "*.onmicrosoft.com") { continue }
            $mx = @(Resolve-DnsSafe -Name $name -Type MX | ForEach-Object { $_.NameExchange }) -join '; '
            $spf = @(Get-TxtRecords -Name $name | Where-Object { $_ -like "v=spf1*" }) -join '; '
            $dmarcTxt = @(Get-TxtRecords -Name "_dmarc.$name" | Where-Object { $_ -like "v=DMARC1*" }) -join '; '
            $mtaSts = @(Get-TxtRecords -Name "_mta-sts.$name" | Where-Object { $_ -like "v=STSv1*" }) -join '; '
            $tlsRpt = @(Get-TxtRecords -Name "_smtp._tls.$name" | Where-Object { $_ -like "v=TLSRPTv1*" }) -join '; '
            $bimi = @(Get-TxtRecords -Name "default._bimi.$name" | Where-Object { $_ -like "v=BIMI1*" }) -join '; '

            $dmarcParts = @{}
            if ($dmarcTxt) {
                foreach ($kv in ($dmarcTxt -split ';')) {
                    if ($kv -match '^\s*(\w+)\s*=\s*(.+?)\s*$') { $dmarcParts[$Matches[1]] = $Matches[2] }
                }
            }
            $dnsRows += [PSCustomObject]@{
                Domain   = $name
                MX       = $mx
                MXPointsToEXO = ($mx -match 'mail\.protection\.outlook\.com')
                SPF      = $spf
                DMARC_p  = $dmarcParts['p'];  DMARC_sp = $dmarcParts['sp']; DMARC_pct = $dmarcParts['pct']
                DMARC_rua = $dmarcParts['rua']; DMARC_ruf = $dmarcParts['ruf']
                DMARC_adkim = $dmarcParts['adkim']; DMARC_aspf = $dmarcParts['aspf']
                'MTA-STS' = $mtaSts
                'TLS-RPT' = $tlsRpt
                BIMI     = $bimi
            }
        }
        Out-Table -CsvName "DomainDns" -Rows $dnsRows -Columns @('Domain','MX','MXPointsToEXO','SPF','DMARC_p','DMARC_sp','DMARC_pct','DMARC_rua','DMARC_ruf','DMARC_adkim','DMARC_aspf','MTA-STS','TLS-RPT','BIMI')

        Add-Line "### DKIM signing configuration"
        Add-Line
        $dkimRows = @($dkim | ForEach-Object {
            [PSCustomObject]@{
                Domain             = $_.Domain
                Enabled            = $_.Enabled
                Selector1CNAME     = $_.Selector1CNAME
                Selector2CNAME     = $_.Selector2CNAME
                KeySize            = $_.KeySize
                LastChecked        = $_.LastChecked
                RotateOnDate       = $_.RotateOnDate
                SelectorBeforeRotateOnDate = $_.SelectorBeforeRotateOnDate
            }
        })
        Out-Table -CsvName "DkimConfig" -Rows $dkimRows -Columns @('Domain','Enabled','Selector1CNAME','Selector2CNAME','KeySize','LastChecked','RotateOnDate','SelectorBeforeRotateOnDate')

        Add-Line "### MOERA (onmicrosoft.com) DKIM/DMARC"
        Add-Line
        if ($script:InitialDomain) {
            $moeraDmarc = @(Get-TxtRecords -Name "_dmarc.$($script:InitialDomain)" | Where-Object { $_ -like "v=DMARC1*" }) -join '; '
            $moeraDkim = @($dkim | Where-Object { $_.Domain -eq $script:InitialDomain } | ForEach-Object { "Enabled=$($_.Enabled); Selector1=$($_.Selector1CNAME); Selector2=$($_.Selector2CNAME)" }) -join '; '
            Out-Table -CsvName "MoeraAuth" -Columns @('Domain','DMARC','DKIM') -Rows @(
                [PSCustomObject]@{ Domain = $script:InitialDomain; DMARC = $moeraDmarc; DKIM = $moeraDkim }
            )
        }
        else {
            Add-Line "_Initial (MOERA) domain not identified._"
            Add-Line
        }

        Add-Line "### ARC trusted sealers"
        Add-Line
        $arc = @()
        try { $arc = @(Get-ArcConfig) } catch { Add-LogEntry -Section 'ARC config' -Reason $_.Exception.Message }
        $arcRows = @($arc | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; ArcTrustedSealers = $_.ArcTrustedSealers }
        })
        Out-Table -CsvName "ArcConfig" -Rows $arcRows -Columns @('Identity','ArcTrustedSealers')
    }
}

function Get-MailFlowSection {
    Invoke-Section -Title $script:SectionTitles.MailFlow -Body {
        Add-Line "### Inbound connectors"
        Add-Line
        $inbound = @()
        try { $inbound = @(Get-InboundConnector) } catch { Add-LogEntry -Section 'Inbound connectors' -Reason $_.Exception.Message }
        $inRows = @($inbound | ForEach-Object {
            $efOn = ($_.EFSkipLastIP -eq $true) -or (@($_.EFSkipIPs).Count -gt 0)
            [PSCustomObject]@{
                Name = $_.Name
                ConnectorType = $_.ConnectorType
                Enabled = $_.Enabled
                SenderDomains = $_.SenderDomains
                SenderIPAddresses = $_.SenderIPAddresses
                RequireTls = $_.RequireTls
                TlsSenderCertificateName = $_.TlsSenderCertificateName
                RestrictDomainsToIPAddresses = $_.RestrictDomainsToIPAddresses
                RestrictDomainsToCertificate = $_.RestrictDomainsToCertificate
                CloudServicesMailEnabled = $_.CloudServicesMailEnabled
                TreatMessagesAsInternal = $_.TreatMessagesAsInternal
                EnhancedFiltering = if ($efOn) { 'On' } else { 'Off' }
                EFSkipLastIP = $_.EFSkipLastIP
                EFSkipIPs = $_.EFSkipIPs
                EFUsers = $_.EFUsers
                EFTestMode = $_.EFTestMode
            }
        })
        Out-Table -CsvName "InboundConnectors" -Rows $inRows -Columns @('Name','ConnectorType','Enabled','SenderDomains','SenderIPAddresses','RequireTls','TlsSenderCertificateName','RestrictDomainsToIPAddresses','RestrictDomainsToCertificate','CloudServicesMailEnabled','TreatMessagesAsInternal','EnhancedFiltering','EFSkipLastIP','EFSkipIPs','EFUsers','EFTestMode')
        $script:Summary['Inbound connectors'] = $inbound.Count

        Add-Line "### Outbound connectors"
        Add-Line
        $outbound = @()
        try { $outbound = @(Get-OutboundConnector) } catch { Add-LogEntry -Section 'Outbound connectors' -Reason $_.Exception.Message }
        $outRows = @($outbound | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                ConnectorType = $_.ConnectorType
                Enabled = $_.Enabled
                RecipientDomains = $_.RecipientDomains
                SmartHosts = $_.SmartHosts
                TlsSettings = $_.TlsSettings
                TlsDomain = $_.TlsDomain
                UseMXRecord = $_.UseMXRecord
                CloudServicesMailEnabled = $_.CloudServicesMailEnabled
                IsTransportRuleScoped = $_.IsTransportRuleScoped
                RouteAllMessagesViaOnPremises = $_.RouteAllMessagesViaOnPremises
            }
        })
        Out-Table -CsvName "OutboundConnectors" -Rows $outRows -Columns @('Name','ConnectorType','Enabled','RecipientDomains','SmartHosts','TlsSettings','TlsDomain','UseMXRecord','CloudServicesMailEnabled','IsTransportRuleScoped','RouteAllMessagesViaOnPremises')
        $script:Summary['Outbound connectors'] = $outbound.Count

        Add-Line "### Transport rules"
        Add-Line
        $rules = @()
        try { $rules = @(Get-TransportRule) } catch { Add-LogEntry -Section 'Transport rules' -Reason $_.Exception.Message }
        $ruleRows = @($rules | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                State = $_.State
                Priority = $_.Priority
                Mode = $_.Mode
                SenderAddressLocation = $_.SenderAddressLocation
                Comments = $_.Comments
                WhenChanged = $_.WhenChanged
            }
        })
        Out-Table -CsvName "TransportRules" -Rows $ruleRows -Columns @('Name','State','Priority','Mode','SenderAddressLocation','Comments','WhenChanged')
        $script:Summary['Transport rules (enabled)'] = @($rules | Where-Object { $_.State -eq 'Enabled' }).Count
        $script:Summary['Transport rules (disabled)'] = @($rules | Where-Object { $_.State -eq 'Disabled' }).Count

        Add-Line "### Remote domains"
        Add-Line
        $remote = @()
        try { $remote = @(Get-RemoteDomain) } catch { Add-LogEntry -Section 'Remote domains' -Reason $_.Exception.Message }
        $remoteRows = @($remote | ForEach-Object {
            [PSCustomObject]@{
                DomainName = $_.DomainName
                AutoForwardEnabled = $_.AutoForwardEnabled
                AutoReplyEnabled = $_.AutoReplyEnabled
                AllowedOOFType = $_.AllowedOOFType
                TNEFEnabled = $_.TNEFEnabled
                CharacterSet = $_.CharacterSet
            }
        })
        Out-Table -CsvName "RemoteDomains" -Rows $remoteRows -Columns @('DomainName','AutoForwardEnabled','AutoReplyEnabled','AllowedOOFType','TNEFEnabled','CharacterSet')

        Add-Line "### Journal rules"
        Add-Line
        $journal = @()
        try { $journal = @(Get-JournalRule) } catch { Add-LogEntry -Section 'Journal rules' -Reason $_.Exception.Message }
        $journalRows = @($journal | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                JournalEmailAddress = $_.JournalEmailAddress
                Scope = $_.Scope
                Recipient = $_.Recipient
                Enabled = $_.Enabled
            }
        })
        Out-Table -CsvName "JournalRules" -Rows $journalRows -Columns @('Name','JournalEmailAddress','Scope','Recipient','Enabled')

        Add-Line "### High Volume Email (HVE) accounts"
        Add-Line
        $hve = @()
        try { $hve = @(Get-MailUser -HVEAccount -ResultSize Unlimited) } catch { Add-LogEntry -Section 'HVE accounts' -Reason $_.Exception.Message }
        if ($hve.Count -eq 0) {
            Add-Line "No HVE accounts in use."
            Add-Line
        }
        else {
            $hveRows = @($hve | ForEach-Object {
                [PSCustomObject]@{ DisplayName = $_.DisplayName; PrimarySmtpAddress = $_.PrimarySmtpAddress; MaxSendPerMinute = $_.MaxSendPerMinute }
            })
            Out-Table -CsvName "HveAccounts" -Rows $hveRows -Columns @('DisplayName','PrimarySmtpAddress','MaxSendPerMinute')
        }
        $script:Summary['HVE accounts'] = $hve.Count
    }

    Invoke-Section -Title "Message trace summary (-IncludeMessageTrace)" -Level 3 -Body {
        if (-not $IncludeMessageTrace) {
            Add-Line "_Skipped - message trace not collected (run with -IncludeMessageTrace)_"
            return
        }
        Assert-Cmdlet Get-MessageTraceV2
        $trace = @(Get-MessageTraceV2 -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -ResultSize 1000)
        if ($trace.Count -eq 0) {
            Add-Line "_No messages returned for the last 7 days (or trace throttled)._"
            return
        }
        Add-Line "Sample of $($trace.Count) messages (last 7 days, up to 1000 rows returned by the first query page)."
        Add-Line
        $statusRows = @($trace | Group-Object Status | ForEach-Object {
            [PSCustomObject]@{ Status = $_.Name; Count = $_.Count }
        })
        Out-Table -CsvName "MessageTraceByStatus" -Rows $statusRows -Columns @('Status','Count')

        Add-Line "#### Messages by direction"
        Add-Line
        $acceptedNames = @($script:AcceptedDomains | ForEach-Object { "$($_.DomainName)".ToLower() })
        $dirCounts = [ordered]@{ Internal = 0; Outbound = 0; Inbound = 0; Other = 0 }
        foreach ($msg in $trace) {
            $dirCounts[(Get-MessageDirection -Sender $msg.SenderAddress -Recipient $msg.RecipientAddress -Domains $acceptedNames)]++
        }
        $dirRows = @($dirCounts.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{ Direction = $_.Key; Count = $_.Value }
        })
        Out-Table -CsvName "MessageTraceByDirection" -Rows $dirRows -Columns @('Direction','Count')

        Register-Csv -Name 'MessageTrace' -Rows $trace
        $script:Summary['Message trace messages sampled'] = $trace.Count
        $script:Summary['Messages sampled (inbound)'] = $dirCounts['Inbound']
        $script:Summary['Messages sampled (outbound)'] = $dirCounts['Outbound']
    }
}

function Get-HybridSection {
    Invoke-Section -Title $script:SectionTitles.Hybrid -Body {
        $rows = @()
        try { $opo = @(Get-OnPremisesOrganization) } catch { $opo = @(); Add-LogEntry -Section 'OnPremisesOrganization' -Reason $_.Exception.Message }
        $rows += @($opo | ForEach-Object {
            [PSCustomObject]@{ Type = 'OnPremisesOrganization'; Name = $_.Name; Details = "HybridDomains=$($_.HybridDomains); Guid=$($_.Guid)" }
        })
        try { $ioc = @(Get-IntraOrganizationConnector) } catch { $ioc = @(); Add-LogEntry -Section 'IntraOrganizationConnector' -Reason $_.Exception.Message }
        $rows += @($ioc | ForEach-Object {
            [PSCustomObject]@{ Type = 'IntraOrganizationConnector'; Name = $_.Name; Details = "Enabled=$($_.Enabled); TargetAddressDomains=$($_.TargetAddressDomains); DiscoveryEndpoint=$($_.DiscoveryEndpoint)" }
        })
        try { $mep = @(Get-MigrationEndpoint) } catch { $mep = @(); Add-LogEntry -Section 'MigrationEndpoint' -Reason $_.Exception.Message }
        $rows += @($mep | ForEach-Object {
            [PSCustomObject]@{ Type = 'MigrationEndpoint'; Name = $_.Identity; Details = "RemoteServer=$($_.RemoteServer); ExchangeVersion=$($_.ExchangeVersion)" }
        })
        try { $batch = @(Get-MigrationBatch) } catch { $batch = @(); Add-LogEntry -Section 'MigrationBatch' -Reason $_.Exception.Message }
        $rows += @($batch | ForEach-Object {
            [PSCustomObject]@{ Type = 'MigrationBatch'; Name = $_.Identity; Details = "Status=$($_.Status); TotalCount=$($_.TotalCount); FinalizedCount=$($_.FinalizedCount)" }
        })
        Out-Table -CsvName "HybridMigration" -Rows $rows -Columns @('Type','Name','Details')
        $script:Summary['Migration batches'] = $batch.Count
    }
}

function Get-SharingSection {
    Invoke-Section -Title $script:SectionTitles.Sharing -Body {
        Add-Line "### Organization relationships"
        Add-Line
        $rel = @()
        try { $rel = @(Get-OrganizationRelationship) } catch { Add-LogEntry -Section 'OrganizationRelationship' -Reason $_.Exception.Message }
        $relRows = @($rel | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                DomainNames = $_.DomainNames
                FreeBusyAccessEnabled = $_.FreeBusyAccessEnabled
                FreeBusyAccessLevel = $_.FreeBusyAccessLevel
                Enabled = $_.Enabled
            }
        })
        Out-Table -CsvName "OrganizationRelationships" -Rows $relRows -Columns @('Name','DomainNames','FreeBusyAccessEnabled','FreeBusyAccessLevel','Enabled')

        Add-Line "### Sharing policies"
        Add-Line
        $sp = @()
        try { $sp = @(Get-SharingPolicy) } catch { Add-LogEntry -Section 'SharingPolicy' -Reason $_.Exception.Message }
        $spRows = @($sp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Domains = $_.Domains
                Default = $_.Default
                Enabled = $_.Enabled
            }
        })
        Out-Table -CsvName "SharingPolicies" -Rows $spRows -Columns @('Name','Domains','Default','Enabled')

        Add-Line "### Availability address spaces"
        Add-Line
        $aas = @()
        try { $aas = @(Get-AvailabilityAddressSpace) } catch { Add-LogEntry -Section 'AvailabilityAddressSpace' -Reason $_.Exception.Message }
        $aasRows = @($aas | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ForestName = $_.ForestName; AccessMethod = $_.AccessMethod }
        })
        Out-Table -CsvName "AvailabilityAddressSpaces" -Rows $aasRows -Columns @('Name','ForestName','AccessMethod')
    }
}

function Get-ClientAccessSection {
    Invoke-Section -Title $script:SectionTitles.ClientAccess -Body {
        Add-Line "### OWA mailbox policies"
        Add-Line
        $owa = @()
        try { $owa = @(Get-OwaMailboxPolicy) } catch { Add-LogEntry -Section 'OwaMailboxPolicy' -Reason $_.Exception.Message }
        $owaRows = @($owa | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                IsDefault = $_.IsDefault
                ClassicAttachmentsEnabled = $_.ClassicAttachmentsEnabled
                ExternalImageProxyEnabled = $_.ExternalImageProxyEnabled
                ThirdPartyFileProvidersEnabled = $_.ThirdPartyFileProvidersEnabled
                ConditionalAccessPolicy = $_.ConditionalAccessPolicy
                AdditionalStorageProvidersAvailable = $_.AdditionalStorageProvidersAvailable
            }
        })
        Out-Table -CsvName "OwaMailboxPolicies" -Rows $owaRows -Columns @('Name','IsDefault','ClassicAttachmentsEnabled','ExternalImageProxyEnabled','ThirdPartyFileProvidersEnabled','ConditionalAccessPolicy','AdditionalStorageProvidersAvailable')

        Add-Line "### ActiveSync organization settings and access rules"
        Add-Line
        $aso = @()
        try { $aso = @(Get-ActiveSyncOrganizationSettings) } catch { Add-LogEntry -Section 'ActiveSyncOrganizationSettings' -Reason $_.Exception.Message }
        $asoRows = @($aso | ForEach-Object {
            [PSCustomObject]@{ DefaultAccessLevel = $_.DefaultAccessLevel; UserMailInsert = $_.UserMailInsert; AdminMailRecipients = $_.AdminMailRecipients }
        })
        Out-Table -CsvName "ActiveSyncOrgSettings" -Rows $asoRows -Columns @('DefaultAccessLevel','UserMailInsert','AdminMailRecipients')
        $asr = @()
        try { $asr = @(Get-ActiveSyncDeviceAccessRule) } catch { Add-LogEntry -Section 'ActiveSyncDeviceAccessRule' -Reason $_.Exception.Message }
        $asrRows = @($asr | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Characteristic = $_.Characteristic; QueryString = $_.QueryString; AccessLevel = $_.AccessLevel }
        })
        Out-Table -CsvName "ActiveSyncDeviceAccessRules" -Rows $asrRows -Columns @('Name','Characteristic','QueryString','AccessLevel')

        Add-Line "### Mobile device mailbox policies"
        Add-Line
        $mdp = @()
        try { $mdp = @(Get-MobileDeviceMailboxPolicy) } catch { Add-LogEntry -Section 'MobileDeviceMailboxPolicy' -Reason $_.Exception.Message }
        $mdpRows = @($mdp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                IsDefault = $_.IsDefault
                PasswordEnabled = $_.PasswordEnabled
                AlphanumericPasswordRequired = $_.AlphanumericPasswordRequired
                MaxInactivityTimeLock = $_.MaxInactivityTimeLock
                AllowNonProvisionableDevices = $_.AllowNonProvisionableDevices
            }
        })
        Out-Table -CsvName "MobileDeviceMailboxPolicies" -Rows $mdpRows -Columns @('Name','IsDefault','PasswordEnabled','AlphanumericPasswordRequired','MaxInactivityTimeLock','AllowNonProvisionableDevices')

        Add-Line "### Protocol usage counts"
        Add-Line
        $cas = @()
        try { $cas = Get-SharedCasMailboxes } catch { Add-LogEntry -Section 'CAS mailboxes' -Reason $_.Exception.Message }
        Out-Table -CsvName "ProtocolUsage" -Columns @('Protocol','Enabled','Disabled') -Rows @(
            [PSCustomObject]@{ Protocol = 'POP'; Enabled = @($cas | Where-Object { $_.PopEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.PopEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'IMAP'; Enabled = @($cas | Where-Object { $_.ImapEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.ImapEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'EWS'; Enabled = @($cas | Where-Object { $_.EwsEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.EwsEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'ActiveSync'; Enabled = @($cas | Where-Object { $_.ActiveSyncEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.ActiveSyncEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'MAPI'; Enabled = @($cas | Where-Object { $_.MapiEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.MapiEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'OWA'; Enabled = @($cas | Where-Object { $_.OWAEnabled }).Count; Disabled = @($cas | Where-Object { -not $_.OWAEnabled }).Count }
            [PSCustomObject]@{ Protocol = 'SMTP AUTH'; Enabled = @($cas | Where-Object { $_.SmtpClientAuthenticationDisabled -eq $false }).Count; Disabled = @($cas | Where-Object { $_.SmtpClientAuthenticationDisabled -eq $true }).Count }
        )

        Add-Line "### Authentication policies"
        Add-Line
        $ap = @()
        try { $ap = @(Get-AuthenticationPolicy) } catch { Add-LogEntry -Section 'AuthenticationPolicy' -Reason $_.Exception.Message }
        $apRows = @($ap | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                AllowBasicAuthPop = $_.AllowBasicAuthPop
                AllowBasicAuthImap = $_.AllowBasicAuthImap
                AllowBasicAuthSmtp = $_.AllowBasicAuthSmtp
                AllowBasicAuthActiveSync = $_.AllowBasicAuthActiveSync
                AllowBasicAuthWebService = $_.AllowBasicAuthWebService
                AllowBasicAuthMapi = $_.AllowBasicAuthMapi
                AllowBasicAuthOfflineAddressBook = $_.AllowBasicAuthOfflineAddressBook
                AllowBasicAuthRpc = $_.AllowBasicAuthRpc
                BlockLegacyAuthActiveSync = $_.BlockLegacyAuthActiveSync
                BlockLegacyAuthImap = $_.BlockLegacyAuthImap
                BlockLegacyAuthMapi = $_.BlockLegacyAuthMapi
                BlockLegacyAuthOfflineAddressBook = $_.BlockLegacyAuthOfflineAddressBook
                BlockLegacyAuthPop = $_.BlockLegacyAuthPop
                BlockLegacyAuthRpc = $_.BlockLegacyAuthRpc
                BlockLegacyAuthWebServices = $_.BlockLegacyAuthWebServices
            }
        })
        Out-Table -CsvName "AuthenticationPolicies" -Rows $apRows -Columns @('Name','AllowBasicAuthPop','AllowBasicAuthImap','AllowBasicAuthSmtp','AllowBasicAuthActiveSync','AllowBasicAuthWebService','AllowBasicAuthMapi','AllowBasicAuthOfflineAddressBook','AllowBasicAuthRpc','BlockLegacyAuthActiveSync','BlockLegacyAuthImap','BlockLegacyAuthMapi','BlockLegacyAuthOfflineAddressBook','BlockLegacyAuthPop','BlockLegacyAuthRpc','BlockLegacyAuthWebServices')
    }
}

function Get-PermissionsSection {
    Invoke-Section -Title $script:SectionTitles.Permissions -Body {
        Add-Line "### Role groups and members"
        Add-Line
        $rg = @()
        try { $rg = @(Get-RoleGroup) } catch { Add-LogEntry -Section 'RoleGroup' -Reason $_.Exception.Message }
        $rgRows = @()
        foreach ($g in $rg) {
            $members = @()
            try { $members = @(Get-RoleGroupMember -Identity $_.Name -ErrorAction Stop) } catch { }
            $rgRows += [PSCustomObject]@{
                RoleGroup = $g.Name
                MemberCount = $members.Count
                Members = @($members | ForEach-Object { $_.Name }) -join '; '
                ManagedBy = $g.ManagedBy
            }
        }
        Out-Table -CsvName "RoleGroups" -Rows $rgRows -Columns @('RoleGroup','MemberCount','Members','ManagedBy')

        Add-Line "### Role assignment policies (user roles)"
        Add-Line
        $rap = @()
        try { $rap = @(Get-RoleAssignmentPolicy) } catch { Add-LogEntry -Section 'RoleAssignmentPolicy' -Reason $_.Exception.Message }
        $mbx = @()
        try { $mbx = Get-SharedMailboxes } catch { }
        $rapRows = @($rap | ForEach-Object {
            $rapName = $_.Name
            $count = @($mbx | Where-Object { $_.RoleAssignmentPolicy -eq $rapName }).Count
            [PSCustomObject]@{
                Name = $rapName
                IsDefault = $_.IsDefault
                AssignedRoles = @($_.AssignedRoles | ForEach-Object { $_.Name }) -join '; '
                MailboxCount = $count
            }
        })
        Out-Table -CsvName "RoleAssignmentPolicies" -Rows $rapRows -Columns @('Name','IsDefault','AssignedRoles','MailboxCount')

        Add-Line "### Custom management roles"
        Add-Line
        $customRoles = @()
        try { $customRoles = @(Get-ManagementRole | Where-Object { -not $_.IsEndUserRole -and -not $_.IsRootRole } ) } catch { Add-LogEntry -Section 'Custom management roles' -Reason $_.Exception.Message }
        $crRows = @($customRoles | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Parent = $_.Parent; RoleType = $_.RoleType }
        })
        Out-Table -CsvName "CustomManagementRoles" -Rows $crRows -Columns @('Name','Parent','RoleType')

        Add-Line "### Direct user role assignments"
        Add-Line
        $assignments = @()
        try { $assignments = @(Get-ManagementRoleAssignment -RoleAssigneeType User) } catch { Add-LogEntry -Section 'Direct role assignments' -Reason $_.Exception.Message }
        $asRows = @($assignments | ForEach-Object {
            [PSCustomObject]@{
                Role = $_.Role
                RoleAssignee = $_.RoleAssigneeName
                AssignmentMethod = $_.AssignmentMethod
                CustomRecipientWriteScope = $_.CustomRecipientWriteScope
                RecipientAdministrativeUnitScope = $_.RecipientAdministrativeUnitScope
            }
        })
        Out-Table -CsvName "DirectRoleAssignments" -Rows $asRows -Columns @('Role','RoleAssignee','AssignmentMethod','CustomRecipientWriteScope','RecipientAdministrativeUnitScope')

        Add-Line "### Management scopes"
        Add-Line
        $scopes = @()
        try { $scopes = @(Get-ManagementScope | Where-Object { -not $_.Exclusive }) } catch { Add-LogEntry -Section 'ManagementScope' -Reason $_.Exception.Message }
        $scRows = @($scopes | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientRestrictionFilter = $_.RecipientRestrictionFilter; Exclusive = $_.Exclusive }
        })
        Out-Table -CsvName "ManagementScopes" -Rows $scRows -Columns @('Name','RecipientRestrictionFilter','Exclusive')

        Add-Line "### RBAC for Applications"
        Add-Line
        $sps = @()
        try { $sps = @(Get-ServicePrincipal) } catch { Add-LogEntry -Section 'ServicePrincipal' -Reason $_.Exception.Message }
        $spAssignments = @()
        try {
            $spAssignments = @(Get-ManagementRoleAssignment -RoleAssigneeType ServicePrincipal -ErrorAction Stop)
        }
        catch {
            # Older modules don't accept the ServicePrincipal assignee type; filter client-side instead.
            try {
                $spIds = @{}
                foreach ($sp in $sps) { $spIds["$($sp.ObjectId)"] = $true; $spIds["$($sp.AppId)"] = $true }
                $spAssignments = @(Get-ManagementRoleAssignment -ErrorAction Stop | Where-Object {
                    $_.RoleAssigneeType -eq 'ServicePrincipal' -or $spIds.ContainsKey("$($_.App)") -or $spIds.ContainsKey("$($_.RoleAssignee)")
                })
                Add-LogEntry -Section 'SP role assignments' -Reason "-RoleAssigneeType ServicePrincipal unsupported; fell back to client-side filtering"
            }
            catch { Add-LogEntry -Section 'SP role assignments' -Reason $_.Exception.Message }
        }
        $spById = @{}
        foreach ($sp in $sps) {
            $spById["$($sp.ObjectId)"] = $sp
            if ($sp.AppId) { $spById["$($sp.AppId)"] = $sp }
            if ($sp.DisplayName) { $spById["$($sp.DisplayName)"] = $sp }
        }
        $rbacRows = @()
        foreach ($a in $spAssignments) {
            $spObj = $null
            foreach ($key in @("$($a.App)", "$($a.RoleAssignee)", "$($a.RoleAssigneeName)")) {
                if ($key -and $spById.ContainsKey($key)) { $spObj = $spById[$key]; break }
            }
            $rbacRows += [PSCustomObject]@{
                DisplayName = if ($spObj) { $spObj.DisplayName } else { $a.RoleAssigneeName }
                AppId = if ($spObj) { $spObj.AppId } else { $null }
                ObjectId = if ($spObj) { $spObj.ObjectId } else { $a.App }
                Role = $a.Role
                CustomResourceScope = $a.CustomResourceScope
                RecipientAdministrativeUnitScope = $a.RecipientAdministrativeUnitScope
            }
        }
        Out-Table -CsvName "RbacForApplications" -Rows $rbacRows -Columns @('DisplayName','AppId','ObjectId','Role','CustomResourceScope','RecipientAdministrativeUnitScope')
        $script:Summary['Service principals'] = $sps.Count
        $script:Summary['App role assignments'] = $rbacRows.Count
    }

    Invoke-Section -Title "Exchange Administrator role holders (-IncludeGraph)" -Level 3 -Body {
        if (-not $IncludeGraph) {
            Add-Line "_Skipped - Graph data not collected (run with -IncludeGraph)_"
            return
        }
        if (-not $script:GraphConnected) {
            Add-Line "_Not available: could not connect to Microsoft Graph - $script:GraphError_"
            Add-LogEntry -Section 'Exchange Administrator role holders' -Reason "Graph not connected: $script:GraphError"
            return
        }
        $role = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" -ErrorAction Stop | Select-Object -First 1
        if (-not $role) {
            Add-Line "_Exchange Administrator role not found or not activated._"
            return
        }
        $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop)
        $rows = @($members | ForEach-Object {
            $o = $_.AdditionalProperties
            [PSCustomObject]@{
                DisplayName = $o['displayName']
                UserPrincipalName = $o['userPrincipalName']
                ObjectType = ($_.AdditionalProperties['@odata.type'] -replace '#microsoft.graph.', '')
            }
        })
        Out-Table -CsvName "ExchangeAdminRoleMembers" -Rows $rows -Columns @('DisplayName','UserPrincipalName','ObjectType')
        $script:Summary['Exchange Administrator role holders'] = $rows.Count
    }
}

function Get-ThreatProtectionSection {
    Invoke-Section -Title $script:SectionTitles.ThreatProtection -Body {
        $defenderAvailable = Test-CmdletAvailable Get-ATPProtectionPolicyRule
        if (-not $defenderAvailable) {
            Add-Line "_Defender for Office 365 not licensed / cmdlet not available - EOP data still collected below._"
            Add-Line
        }

        Add-Line "### Anti-phishing policies"
        Add-Line
        $phish = @()
        try { $phish = @(Get-AntiPhishPolicy) } catch { Add-LogEntry -Section 'AntiPhishPolicy' -Reason $_.Exception.Message }
        $phishRules = @()
        try { $phishRules = @(Get-AntiPhishRule) } catch { }
        $pRows = @($phish | ForEach-Object {
            $r = @($phishRules | Where-Object { $_.AntiPhishPolicy -eq $_.Name }) | Select-Object -First 1
            [PSCustomObject]@{
                Name = $_.Name
                Enabled = $_.Enabled
                RuleState = $r.State
                RulePriority = $r.Priority
                ImpersonationAction = $_.TargetedUserProtectionAction
                EnableTargetedUserProtection = $_.EnableTargetedUserProtection
                EnableOrganizationDomainsProtection = $_.EnableOrganizationDomainsProtection
                EnableMailboxIntelligence = $_.EnableMailboxIntelligence
                EnableMailboxIntelligenceProtection = $_.EnableMailboxIntelligenceProtection
                EnableSpoofIntelligence = $_.EnableSpoofIntelligence
                AuthenticationFailAction = $_.AuthenticationFailAction
                HonorDmarcPolicy = $_.HonorDmarcPolicy
                PhishThresholdLevel = $_.PhishThresholdLevel
                TargetedDomainProtectionAction = $_.TargetedDomainProtectionAction
                IncludedUsers = $r.SentTo
                IncludedGroups = $r.SentToMemberOf
                IncludedDomains = $r.RecipientDomainIs
            }
        })
        Out-Table -CsvName "AntiPhishPolicies" -Rows $pRows -Columns @('Name','Enabled','RuleState','RulePriority','EnableTargetedUserProtection','ImpersonationAction','EnableOrganizationDomainsProtection','EnableMailboxIntelligence','EnableMailboxIntelligenceProtection','EnableSpoofIntelligence','AuthenticationFailAction','HonorDmarcPolicy','PhishThresholdLevel','TargetedDomainProtectionAction','IncludedUsers','IncludedGroups','IncludedDomains')

        Add-Line "### Anti-spam inbound policies"
        Add-Line
        $spamIn = @()
        try { $spamIn = @(Get-HostedContentFilterPolicy) } catch { Add-LogEntry -Section 'HostedContentFilterPolicy' -Reason $_.Exception.Message }
        $spamInRules = @()
        try { $spamInRules = @(Get-HostedContentFilterRule) } catch { }
        $siRows = @($spamIn | ForEach-Object {
            $r = @($spamInRules | Where-Object { $_.HostedContentFilterPolicy -eq $_.Name }) | Select-Object -First 1
            [PSCustomObject]@{
                Name = $_.Name
                Enabled = $_.Enabled
                RuleState = $r.State
                RulePriority = $r.Priority
                SpamAction = $_.SpamAction
                HighConfidenceSpamAction = $_.HighConfidenceSpamAction
                PhishSpamAction = $_.PhishSpamAction
                HighConfidencePhishAction = $_.HighConfidencePhishAction
                BulkSpamAction = $_.BulkSpamAction
                BulkThreshold = $_.BulkThreshold
                IncreaseScoreWithImageLinks = $_.IncreaseScoreWithImageLinks
                AllowedSenders = $_.AllowedSenders
                AllowedSenderDomains = $_.AllowedSenderDomains
                BlockedSenders = $_.BlockedSenders
                BlockedSenderDomains = $_.BlockedSenderDomains
                QuarantineRetentionPeriod = $_.QuarantineRetentionPeriod
                SpamQuarantineTag = $_.SpamQuarantineTag
                HighConfidenceSpamQuarantineTag = $_.HighConfidenceSpamQuarantineTag
                PhishQuarantineTag = $_.PhishQuarantineTag
                HighConfidencePhishQuarantineTag = $_.HighConfidencePhishQuarantineTag
                BulkQuarantineTag = $_.BulkQuarantineTag
            }
        })
        Out-Table -CsvName "AntiSpamInboundPolicies" -Rows $siRows -Columns @('Name','Enabled','RuleState','RulePriority','SpamAction','HighConfidenceSpamAction','PhishSpamAction','HighConfidencePhishAction','BulkSpamAction','BulkThreshold','AllowedSenders','AllowedSenderDomains','BlockedSenders','BlockedSenderDomains','QuarantineRetentionPeriod','SpamQuarantineTag','HighConfidenceSpamQuarantineTag','PhishQuarantineTag','HighConfidencePhishQuarantineTag','BulkQuarantineTag')

        Add-Line "### Connection filter policy"
        Add-Line
        $cf = @()
        try { $cf = @(Get-HostedConnectionFilterPolicy) } catch { Add-LogEntry -Section 'HostedConnectionFilterPolicy' -Reason $_.Exception.Message }
        $cfRows = @($cf | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Enabled = $_.Enabled
                IPAllowList = $_.IPAllowList
                IPBlockList = $_.IPBlockList
                EnableSafeList = $_.EnableSafeList
            }
        })
        Out-Table -CsvName "ConnectionFilterPolicy" -Rows $cfRows -Columns @('Name','Enabled','IPAllowList','IPBlockList','EnableSafeList')

        Add-Line "### Outbound spam policies"
        Add-Line
        $spamOut = @()
        try { $spamOut = @(Get-HostedOutboundSpamFilterPolicy) } catch { Add-LogEntry -Section 'HostedOutboundSpamFilterPolicy' -Reason $_.Exception.Message }
        $soRows = @($spamOut | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Enabled = $_.Enabled
                RecipientLimitExternalPerHour = $_.RecipientLimitExternalPerHour
                RecipientLimitInternalPerHour = $_.RecipientLimitInternalPerHour
                RecipientLimitPerDay = $_.RecipientLimitPerDay
                AutoForwardingMode = $_.AutoForwardingMode
                NotifyOutboundSpam = $_.NotifyOutboundSpam
                NotifyOutboundSpamRecipients = $_.NotifyOutboundSpamRecipients
            }
        })
        Out-Table -CsvName "OutboundSpamPolicies" -Rows $soRows -Columns @('Name','Enabled','RecipientLimitExternalPerHour','RecipientLimitInternalPerHour','RecipientLimitPerDay','AutoForwardingMode','NotifyOutboundSpam','NotifyOutboundSpamRecipients')

        Add-Line "### Anti-malware policies"
        Add-Line
        $mal = @()
        try { $mal = @(Get-MalwareFilterPolicy) } catch { Add-LogEntry -Section 'MalwareFilterPolicy' -Reason $_.Exception.Message }
        $malRules = @()
        try { $malRules = @(Get-MalwareFilterRule) } catch { }
        $malRows = @($mal | ForEach-Object {
            $r = @($malRules | Where-Object { $_.MalwareFilterPolicy -eq $_.Name }) | Select-Object -First 1
            [PSCustomObject]@{
                Name = $_.Name
                Enabled = $_.Enabled
                RuleState = $r.State
                RulePriority = $r.Priority
                EnableFileFilter = $_.EnableFileFilter
                FileTypeCount = @($_.FileTypes).Count
                FileTypes = $_.FileTypes
                ZapEnabled = $_.ZapEnabled
                EnableInternalSenderAdminNotifications = $_.EnableInternalSenderAdminNotifications
                InternalSenderAdminAddress = $_.InternalSenderAdminAddress
                EnableExternalSenderAdminNotifications = $_.EnableExternalSenderAdminNotifications
                ExternalSenderAdminAddress = $_.ExternalSenderAdminAddress
                Action = $_.Action
            }
        })
        Out-Table -CsvName "MalwareFilterPolicies" -Rows $malRows -Columns @('Name','Enabled','RuleState','RulePriority','EnableFileFilter','FileTypeCount','FileTypes','ZapEnabled','EnableInternalSenderAdminNotifications','InternalSenderAdminAddress','EnableExternalSenderAdminNotifications','ExternalSenderAdminAddress','Action')

        Add-Line "### Safe Attachments policies"
        Add-Line
        $sa = @()
        try { $sa = @(Get-SafeAttachmentPolicy) } catch { Add-LogEntry -Section 'SafeAttachmentPolicy' -Reason $_.Exception.Message }
        $saRules = @()
        try { $saRules = @(Get-SafeAttachmentRule) } catch { }
        $saRows = @($sa | ForEach-Object {
            $r = @($saRules | Where-Object { $_.SafeAttachmentPolicy -eq $_.Name }) | Select-Object -First 1
            [PSCustomObject]@{
                Name = $_.Name
                Enable = $_.Enable
                RuleState = $r.State
                RulePriority = $r.Priority
                Action = $_.Action
                Redirect = $_.Redirect
                RedirectAddress = $_.RedirectAddress
            }
        })
        Out-Table -CsvName "SafeAttachmentPolicies" -Rows $saRows -Columns @('Name','Enable','RuleState','RulePriority','Action','Redirect','RedirectAddress')

        Add-Line "### ATP policy for O365"
        Add-Line
        $atp = @()
        try { $atp = @(Get-AtpPolicyForO365) } catch { Add-LogEntry -Section 'AtpPolicyForO365' -Reason $_.Exception.Message }
        $atpRows = @($atp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                EnableATPForSPOTeamsODB = $_.EnableATPForSPOTeamsODB
                EnableSafeDocs = $_.EnableSafeDocs
                AllowSafeDocsOpen = $_.AllowSafeDocsOpen
                EnableSafeLinksForO365Clients = $_.EnableSafeLinksForO365Clients
                TrackClicks = $_.TrackClicks
                AllowClickThrough = $_.AllowClickThrough
            }
        })
        Out-Table -CsvName "AtpPolicyForO365" -Rows $atpRows -Columns @('Name','EnableATPForSPOTeamsODB','EnableSafeDocs','AllowSafeDocsOpen','EnableSafeLinksForO365Clients','TrackClicks','AllowClickThrough')

        Add-Line "### Safe Links policies"
        Add-Line
        $sl = @()
        try { $sl = @(Get-SafeLinksPolicy) } catch { Add-LogEntry -Section 'SafeLinksPolicy' -Reason $_.Exception.Message }
        $slRules = @()
        try { $slRules = @(Get-SafeLinksRule) } catch { }
        $slRows = @($sl | ForEach-Object {
            $r = @($slRules | Where-Object { $_.SafeLinksPolicy -eq $_.Name }) | Select-Object -First 1
            [PSCustomObject]@{
                Name = $_.Name
                IsEnabled = $_.IsEnabled
                RuleState = $r.State
                RulePriority = $r.Priority
                ScanUrl = $_.ScanUrl
                EnableSafeLinksForEmail = $_.EnableSafeLinksForEmail
                EnableSafeLinksForTeams = $_.EnableSafeLinksForTeams
                EnableSafeLinksForOffice = $_.EnableSafeLinksForOffice
                EnableForInternalSenders = $_.EnableForInternalSenders
                TrackClicks = $_.TrackClicks
                AllowClickThrough = $_.AllowClickThrough
                DoNotRewriteUrls = $_.DoNotRewriteUrls
                DeliverMessageAfterScan = $_.DeliverMessageAfterScan
                DisableUrlRewrite = $_.DisableUrlRewrite
            }
        })
        Out-Table -CsvName "SafeLinksPolicies" -Rows $slRows -Columns @('Name','IsEnabled','RuleState','RulePriority','ScanUrl','EnableSafeLinksForEmail','EnableSafeLinksForTeams','EnableSafeLinksForOffice','EnableForInternalSenders','TrackClicks','AllowClickThrough','DoNotRewriteUrls','DeliverMessageAfterScan','DisableUrlRewrite')

        Add-Line "### Preset security policies"
        Add-Line
        $presetRows = @()
        try {
            $eop = @(Get-EOPProtectionPolicyRule)
            $presetRows += @($eop | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    State = $_.State
                    Priority = $_.Priority
                    Type = 'EOP'
                    SentTo = $_.SentTo
                    SentToMemberOf = $_.SentToMemberOf
                    RecipientDomainIs = $_.RecipientDomainIs
                }
            })
        } catch { Add-LogEntry -Section 'EOPProtectionPolicyRule' -Reason $_.Exception.Message }
        try {
            $atpPreset = @(Get-ATPProtectionPolicyRule)
            $presetRows += @($atpPreset | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_.Name
                    State = $_.State
                    Priority = $_.Priority
                    Type = 'ATP'
                    SentTo = $_.SentTo
                    SentToMemberOf = $_.SentToMemberOf
                    RecipientDomainIs = $_.RecipientDomainIs
                }
            })
        } catch { Add-LogEntry -Section 'ATPProtectionPolicyRule' -Reason $_.Exception.Message }
        Out-Table -CsvName "PresetSecurityPolicies" -Rows $presetRows -Columns @('Name','State','Priority','Type','SentTo','SentToMemberOf','RecipientDomainIs')

        Add-Line "### Built-in protection rule"
        Add-Line
        $builtin = @()
        try { $builtin = @(Get-ATPBuiltInProtectionRule) } catch { Add-LogEntry -Section 'ATPBuiltInProtectionRule' -Reason $_.Exception.Message }
        $biRows = @($builtin | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; State = $_.State; SentTo = $_.SentTo; SentToMemberOf = $_.SentToMemberOf; RecipientDomainIs = $_.RecipientDomainIs }
        })
        Out-Table -CsvName "BuiltInProtectionRule" -Rows $biRows -Columns @('Name','State','SentTo','SentToMemberOf','RecipientDomainIs')

        Add-Line "### Tenant Allow/Block List"
        Add-Line
        $tabl = @()
        foreach ($listType in @('Sender','Url','FileHash','IP')) {
            try { $tabl += @(Get-TenantAllowBlockListItems -ListType $listType -ErrorAction Stop) }
            catch { Add-LogEntry -Section "TenantAllowBlockListItems ($listType)" -Reason $_.Exception.Message }
        }
        $tablRows = @($tabl | ForEach-Object {
            [PSCustomObject]@{
                ListType = $_.ListType
                Value = $_.Value
                Action = $_.Action
                ExpirationDate = $_.ExpirationDate
                Notes = $_.Notes
            }
        })
        Out-Table -CsvName "TenantAllowBlockList" -Rows $tablRows -Columns @('ListType','Value','Action','ExpirationDate','Notes')
        $spoof = @()
        try { $spoof = @(Get-TenantAllowBlockListSpoofItems) } catch { Add-LogEntry -Section 'TenantAllowBlockListSpoofItems' -Reason $_.Exception.Message }
        $spoofRows = @($spoof | ForEach-Object {
            [PSCustomObject]@{ SpoofedUser = $_.SpoofedUser; SendingInfrastructure = $_.SendingInfrastructure; Action = $_.Action; SpoofType = $_.SpoofType }
        })
        Out-Table -CsvName "TenantAllowBlockListSpoof" -Rows $spoofRows -Columns @('SpoofedUser','SendingInfrastructure','Action','SpoofType')

        Add-Line "### Advanced delivery (SecOps mailboxes and phishing simulations)"
        Add-Line
        $secOps = @()
        try { $secOps = @(Get-SecOpsOverridePolicy) } catch { Add-LogEntry -Section 'SecOpsOverridePolicy' -Reason $_.Exception.Message }
        $secOpsRules = @()
        try { $secOpsRules = @(Get-ExoSecOpsOverrideRule) } catch { }
        $secOpsRows = @($secOps | ForEach-Object {
            $r = @($secOpsRules | Where-Object { $_.Policy -eq $_.Identity }) | Select-Object -First 1
            [PSCustomObject]@{ Name = $_.Name; SentTo = $r.SentTo; SentToMemberOf = $r.SentToMemberOf }
        })
        Out-Table -CsvName "SecOpsOverridePolicy" -Rows $secOpsRows -Columns @('Name','SentTo','SentToMemberOf')
        $phishSim = @()
        try { $phishSim = @(Get-PhishSimOverridePolicy) } catch { Add-LogEntry -Section 'PhishSimOverridePolicy' -Reason $_.Exception.Message }
        $phishSimRules = @()
        try { $phishSimRules = @(Get-ExoPhishSimOverrideRule) } catch { }
        $simRows = @($phishSim | ForEach-Object {
            $r = @($phishSimRules | Where-Object { $_.Policy -eq $_.Identity }) | Select-Object -First 1
            [PSCustomObject]@{ Name = $_.Name; Domains = $r.Domains; SenderIpRanges = $r.SenderIpRanges }
        })
        Out-Table -CsvName "PhishSimOverridePolicy" -Rows $simRows -Columns @('Name','Domains','SenderIpRanges')
        $simUrls = @()
        try { $simUrls = @(Get-TenantAllowBlockListItems -ListType Url -ListSubType AdvancedDelivery) } catch { Add-LogEntry -Section 'Simulation URLs' -Reason $_.Exception.Message }
        $simUrlRows = @($simUrls | ForEach-Object {
            [PSCustomObject]@{ Value = $_.Value; ExpirationDate = $_.ExpirationDate }
        })
        Out-Table -CsvName "SimulationUrls" -Rows $simUrlRows -Columns @('Value','ExpirationDate')

        Add-Line "### Quarantine policies"
        Add-Line
        $qp = @()
        try { $qp = @(Get-QuarantinePolicy) } catch { Add-LogEntry -Section 'QuarantinePolicy' -Reason $_.Exception.Message }
        $qpRows = @($qp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                EndUserQuarantinePermissionsValue = $_.EndUserQuarantinePermissionsValue
                ESNEnabled = $_.ESNEnabled
                QuarantineRetentionDays = $_.QuarantineRetentionDays
            }
        })
        Out-Table -CsvName "QuarantinePolicies" -Rows $qpRows -Columns @('Name','EndUserQuarantinePermissionsValue','ESNEnabled','QuarantineRetentionDays')
        $qgs = @()
        try { $qgs = @(Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy) } catch { }
        $qgRows = @($qgs | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; EndUserSpamNotificationFrequency = $_.EndUserSpamNotificationFrequency; ESNEnabled = $_.ESNEnabled }
        })
        Out-Table -CsvName "GlobalQuarantineSettings" -Rows $qgRows -Columns @('Name','EndUserSpamNotificationFrequency','ESNEnabled')

        Add-Line "### Priority account protection and reporting"
        Add-Line
        $ets = @()
        try { $ets = @(Get-EmailTenantSettings) } catch { Add-LogEntry -Section 'EmailTenantSettings' -Reason $_.Exception.Message }
        $etsRows = @($ets | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; EnablePriorityAccountProtection = $_.EnablePriorityAccountProtection }
        })
        Out-Table -CsvName "EmailTenantSettings" -Rows $etsRows -Columns @('Identity','EnablePriorityAccountProtection')
        $rsp = @()
        try { $rsp = @(Get-ReportSubmissionPolicy) } catch { Add-LogEntry -Section 'ReportSubmissionPolicy' -Reason $_.Exception.Message }
        $rspRows = @($rsp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                EnableReportToMicrosoft = $_.EnableReportToMicrosoft
                ReportChatMessageEnabled = $_.ReportChatMessageEnabled
                ReportJunkToCustomizedAddress = $_.ReportJunkToCustomizedAddress
                ReportNotJunkToCustomizedAddress = $_.ReportNotJunkToCustomizedAddress
                ReportPhishToCustomizedAddress = $_.ReportPhishToCustomizedAddress
                ReportJunkAddresses = $_.ReportJunkAddresses
                ReportPhishAddresses = $_.ReportPhishAddresses
            }
        })
        Out-Table -CsvName "ReportSubmissionPolicy" -Rows $rspRows -Columns @('Name','EnableReportToMicrosoft','ReportChatMessageEnabled','ReportJunkToCustomizedAddress','ReportNotJunkToCustomizedAddress','ReportPhishToCustomizedAddress','ReportJunkAddresses','ReportPhishAddresses')
        $tpp = @()
        try { $tpp = @(Get-TeamsProtectionPolicy) } catch { Add-LogEntry -Section 'TeamsProtectionPolicy' -Reason $_.Exception.Message }
        $tppRows = @($tpp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ZAPForTeamsEnabled = $_.ZAPForTeamsEnabled; ZapEnabled = $_.ZapEnabled }
        })
        Out-Table -CsvName "TeamsProtectionPolicy" -Rows $tppRows -Columns @('Name','ZAPForTeamsEnabled','ZapEnabled')
    }
}

function Get-AlertPoliciesSection {
    Invoke-Section -Title $script:SectionTitles.AlertPolicies -Body {
        if (Add-PurviewStatusOrThrow -Section 'Alert policies') { return }
        Assert-Cmdlet Get-ProtectionAlert
        $alerts = @(Get-ProtectionAlert)
        $rows = @($alerts | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Category = $_.Category
                Severity = $_.Severity
                Disabled = $_.Disabled
                IsSystemRule = $_.IsSystemRule
                NotifyUser = $_.NotifyUser
                ThreatType = $_.ThreatType
            }
        })
        Out-Table -CsvName "ProtectionAlerts" -Rows $rows -Columns @('Name','Category','Severity','Disabled','IsSystemRule','NotifyUser','ThreatType')
        $script:Summary['Alert policies'] = $alerts.Count
    }
}

function Get-ComplianceSection {
    Invoke-Section -Title $script:SectionTitles.Compliance -Body {
        Add-Line "### MRM retention policies and tags"
        Add-Line
        $rp = @()
        try { $rp = @(Get-RetentionPolicy) } catch { Add-LogEntry -Section 'RetentionPolicy' -Reason $_.Exception.Message }
        $rpRows = @($rp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                RetentionPolicyTagLinks = @($_.RetentionPolicyTagLinks | ForEach-Object { $_.Name }) -join '; '
                IsDefault = $_.IsDefault
            }
        })
        Out-Table -CsvName "RetentionPolicies" -Rows $rpRows -Columns @('Name','RetentionPolicyTagLinks','IsDefault')
        $tags = @()
        try { $tags = @(Get-RetentionPolicyTag) } catch { Add-LogEntry -Section 'RetentionPolicyTag' -Reason $_.Exception.Message }
        $tagRows = @($tags | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Type = $_.Type
                RetentionAction = $_.RetentionAction
                AgeLimitForRetention = $_.AgeLimitForRetention
                RetentionEnabled = $_.RetentionEnabled
                IsDefaultModeratedRecoveryPolicyTag = $_.IsDefaultModeratedRecoveryPolicyTag
            }
        })
        Out-Table -CsvName "RetentionPolicyTags" -Rows $tagRows -Columns @('Name','Type','RetentionAction','AgeLimitForRetention','RetentionEnabled','IsDefaultModeratedRecoveryPolicyTag')

        Add-Line "### IRM and OME configuration"
        Add-Line
        $irm = @()
        try { $irm = @(Get-IRMConfiguration) } catch { Add-LogEntry -Section 'IRMConfiguration' -Reason $_.Exception.Message }
        $irmRows = @($irm | ForEach-Object {
            [PSCustomObject]@{
                Identity = $_.Identity
                InternalLicensingEnabled = $_.InternalLicensingEnabled
                ExternalLicensingEnabled = $_.ExternalLicensingEnabled
                AzureRMSLicensingEnabled = $_.AzureRMSLicensingEnabled
                TransportDecryptionSetting = $_.TransportDecryptionSetting
                JournalReportDecryptionEnabled = $_.JournalReportDecryptionEnabled
                SearchEnabled = $_.SearchEnabled
            }
        })
        Out-Table -CsvName "IrmConfiguration" -Rows $irmRows -Columns @('Identity','InternalLicensingEnabled','ExternalLicensingEnabled','AzureRMSLicensingEnabled','TransportDecryptionSetting','JournalReportDecryptionEnabled','SearchEnabled')
        $ome = @()
        try { $ome = @(Get-OMEConfiguration) } catch { Add-LogEntry -Section 'OMEConfiguration' -Reason $_.Exception.Message }
        $omeRows = @($ome | ForEach-Object {
            [PSCustomObject]@{
                Identity = $_.Identity
                OTPEnabled = $_.OTPEnabled
                SocialIdSignIn = $_.SocialIdSignIn
                ExternalMailExpiryInDays = $_.ExternalMailExpiryInDays
            }
        })
        Out-Table -CsvName "OmeConfiguration" -Rows $omeRows -Columns @('Identity','OTPEnabled','SocialIdSignIn','ExternalMailExpiryInDays')
    }

    Invoke-Section -Title "Purview retention policies covering Exchange" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Purview retention') { return }
        Assert-Cmdlet Get-RetentionCompliancePolicy
        $policies = @(Get-RetentionCompliancePolicy -DistributionDetail)
        $exo = @($policies | Where-Object {
            @($_.ExchangeLocation).Count -gt 0 -and "$($_.ExchangeLocation)" -notmatch '^\s*$'
        })
        $rows = @($exo | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Mode = $_.Mode
                Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation
                ExchangeLocationException = $_.ExchangeLocationException
                Rules = @(@(Get-RetentionComplianceRule -Policy $_.Name -ErrorAction SilentlyContinue) | ForEach-Object { $_.Name }) -join '; '
            }
        })
        Out-Table -CsvName "PurviewRetentionPolicies" -Rows $rows -Columns @('Name','Mode','Enabled','ExchangeLocation','ExchangeLocationException','Rules')
        $script:Summary['Purview retention policies (Exchange)'] = $exo.Count
    }

    Invoke-Section -Title "DLP policies covering Exchange" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'DLP policies') { return }
        Assert-Cmdlet Get-DlpCompliancePolicy
        $policies = @(Get-DlpCompliancePolicy)
        $exo = @($policies | Where-Object { @($_.ExchangeLocation).Count -gt 0 })
        $rows = @($exo | ForEach-Object {
            $rules = @()
            try { $rules = @(Get-DlpComplianceRule -Policy $_.Name -ErrorAction Stop) } catch { }
            [PSCustomObject]@{
                Name = $_.Name
                Mode = $_.Mode
                Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation
                Rules = @($rules | ForEach-Object { $_.Name }) -join '; '
            }
        })
        Out-Table -CsvName "DlpPolicies" -Rows $rows -Columns @('Name','Mode','Enabled','ExchangeLocation','Rules')
        $script:Summary['DLP policies (Exchange)'] = $exo.Count
    }

    Invoke-Section -Title "Sensitivity labels and label policies" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Sensitivity labels') { return }
        Assert-Cmdlet Get-Label
        $labels = @(Get-Label)
        $labelRows = @($labels | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                DisplayName = $_.DisplayName
                ContentType = $_.ContentType
                Disabled = $_.Disabled
                Priority = $_.Priority
                Tooltip = $_.Tooltip
            }
        })
        Out-Table -CsvName "SensitivityLabels" -Rows $labelRows -Columns @('Name','DisplayName','ContentType','Disabled','Priority','Tooltip')
        $script:Summary['Sensitivity labels'] = $labels.Count

        $labelPolicies = @()
        try { $labelPolicies = @(Get-LabelPolicy) } catch { Add-LogEntry -Section 'LabelPolicy' -Reason $_.Exception.Message }
        $lpRows = @($labelPolicies | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Labels = @($_.Labels) -join '; '
                ExchangeLocation = $_.ExchangeLocation
                Enabled = $_.Enabled
            }
        })
        Out-Table -CsvName "LabelPolicies" -Rows $lpRows -Columns @('Name','Labels','ExchangeLocation','Enabled')
    }

    Invoke-Section -Title "Retention labels" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Retention labels') { return }
        Assert-Cmdlet Get-ComplianceTag
        $tags = @(Get-ComplianceTag)
        $rows = @($tags | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                RetentionAction = $_.RetentionAction
                RetentionDuration = $_.RetentionDuration
                IsRecordLabel = $_.IsRecordLabel
                EventType = $_.EventType
                Notes = $_.Notes
            }
        })
        Out-Table -CsvName "ComplianceTags" -Rows $rows -Columns @('Name','RetentionAction','RetentionDuration','IsRecordLabel','EventType','Notes')
        $script:Summary['Retention labels'] = $tags.Count
    }
}

function Get-GraphSection {
    Invoke-Section -Title $script:SectionTitles.Graph -Body {
        if (-not $IncludeGraph) {
            Add-Line "_Skipped - Graph data not collected (run with -IncludeGraph)_"
            return
        }
        if (-not $script:GraphConnected) {
            Add-Line "_Not available: could not connect to Microsoft Graph - $script:GraphError_"
            Add-LogEntry -Section 'Graph' -Reason "Graph not connected: $script:GraphError"
            return
        }
        Add-Line "### Secure Score (latest)"
        Add-Line
        $score = Get-MgSecuritySecureScore -Top 1 -ErrorAction Stop | Select-Object -First 1
        if ($score) {
            Out-Table -CsvName "SecureScore" -Columns @('CreatedDateTime','CurrentScore','MaxScore') -Rows @(
                [PSCustomObject]@{
                    CreatedDateTime = $score.CreatedDateTime
                    CurrentScore = $score.CurrentScore
                    MaxScore = $score.MaxScore
                }
            )
            Add-Line "### Exchange-related Secure Score controls"
            Add-Line
            $ctrlRows = @($score.ControlScores | Where-Object {
                $_.ControlName -match 'Exchange|EXO|Mail|Outlook|Spam|Phish|DKIM|DMARC|SPF|Forwarding' -or
                $_.Description -match 'Exchange|mail|phish|spam|forwarding|DKIM|DMARC|SPF'
            } | ForEach-Object {
                [PSCustomObject]@{
                    ControlName = $_.ControlName
                    Score = $_.Score
                    ScoreInPercentage = $_.ScoreInPercentage
                    ControlCategory = $_.ControlCategory
                    ImplementationStatus = $_.ImplementationStatus
                }
            })
            Out-Table -CsvName "SecureScoreExchangeControls" -Rows $ctrlRows -Columns @('ControlName','Score','ScoreInPercentage','ControlCategory','ImplementationStatus')
        }
        else {
            Add-Line "_No Secure Score data returned._"
        }
    }
}

function Write-Appendix {
    Add-Line ("## " + $script:SectionTitles.Appendix)
    Add-Line
    if ($script:CollectionLog.Count -eq 0) {
        Add-Line "No skipped or failed sections."
    }
    else {
        $logRows = @($script:CollectionLog)
        Add-Line (ConvertTo-MdTable -Rows $logRows -Columns @('Section','Reason','Time'))
    }
    Add-Line
    Add-Line "### Parameters used"
    Add-Line
    $paramRows = @(
        [PSCustomObject]@{ Parameter = 'CustomerName'; Value = $CustomerName }
        [PSCustomObject]@{ Parameter = 'ReviewedBy'; Value = $ReviewedBy }
        [PSCustomObject]@{ Parameter = 'ReviewDate'; Value = $ReviewDate.ToString('yyyy-MM-dd') }
        [PSCustomObject]@{ Parameter = 'OutputPath'; Value = $OutputPath }
        [PSCustomObject]@{ Parameter = 'UserPrincipalName'; Value = $UserPrincipalName }
        [PSCustomObject]@{ Parameter = 'AppId'; Value = if ($AppId) { $AppId } else { '(interactive)' } }
        [PSCustomObject]@{ Parameter = 'IncludePurview'; Value = [bool]$IncludePurview }
        [PSCustomObject]@{ Parameter = 'IncludeMailboxStatistics'; Value = [bool]$IncludeMailboxStatistics }
        [PSCustomObject]@{ Parameter = 'IncludeMessageTrace'; Value = [bool]$IncludeMessageTrace }
        [PSCustomObject]@{ Parameter = 'IncludeGraph'; Value = [bool]$IncludeGraph }
        [PSCustomObject]@{ Parameter = 'ExportCsv'; Value = [bool]$ExportCsv }
        [PSCustomObject]@{ Parameter = 'MaxRows'; Value = $MaxRows }
        [PSCustomObject]@{ Parameter = 'DnsServer'; Value = if ($DnsServer) { $DnsServer } else { '(default)' } }
    )
    Add-Line (ConvertTo-MdTable -Rows $paramRows -Columns @('Parameter','Value'))
    Add-Line
    $duration = (Get-Date) - $script:RunStart
    Add-Line "### Run duration"
    Add-Line
    Add-Line "$([math]::Round($duration.TotalMinutes, 1)) minutes ($($duration.ToString('hh\:mm\:ss')))"
    Add-Line
    Add-Line "### Cmdlet reference"
    Add-Line
    $cmdletList = @(
        'EXO/EOP: Get-OrganizationConfig, Get-ExternalInOutlook, Get-TransportConfig, Get-AdminAuditLogConfig,',
        'Get-EXOMailbox, Get-EXOCASMailbox, Get-EXOMailboxStatistics, Get-MailboxAuditBypassAssociation, Get-User,',
        'Get-MailUser (-HVEAccount), Get-MailContact, Get-DistributionGroup, Get-DynamicDistributionGroup, Get-UnifiedGroup,',
        'Get-AcceptedDomain, Get-DkimSigningConfig, Get-ArcConfig, Resolve-DnsName, Get-InboundConnector, Get-OutboundConnector,',
        'Get-TransportRule, Get-RemoteDomain, Get-JournalRule, Get-MessageTraceV2, Get-OnPremisesOrganization,',
        'Get-IntraOrganizationConnector, Get-MigrationEndpoint, Get-MigrationBatch, Get-OrganizationRelationship,',
        'Get-SharingPolicy, Get-AvailabilityAddressSpace, Get-OwaMailboxPolicy, Get-ActiveSyncOrganizationSettings,',
        'Get-ActiveSyncDeviceAccessRule, Get-MobileDeviceMailboxPolicy, Get-AuthenticationPolicy, Get-RoleGroup,',
        'Get-RoleGroupMember, Get-RoleAssignmentPolicy, Get-ManagementRole, Get-ManagementRoleAssignment, Get-ManagementScope,',
        'Get-ServicePrincipal, Get-AntiPhishPolicy/Rule, Get-HostedContentFilterPolicy/Rule, Get-HostedConnectionFilterPolicy,',
        'Get-HostedOutboundSpamFilterPolicy, Get-MalwareFilterPolicy/Rule, Get-SafeAttachmentPolicy/Rule, Get-AtpPolicyForO365,',
        'Get-SafeLinksPolicy/Rule, Get-EOPProtectionPolicyRule, Get-ATPProtectionPolicyRule, Get-ATPBuiltInProtectionRule,',
        'Get-TenantAllowBlockListItems, Get-TenantAllowBlockListSpoofItems, Get-SecOpsOverridePolicy, Get-ExoSecOpsOverrideRule,',
        'Get-PhishSimOverridePolicy, Get-ExoPhishSimOverrideRule, Get-QuarantinePolicy, Get-EmailTenantSettings,',
        'Get-ReportSubmissionPolicy, Get-TeamsProtectionPolicy, Get-RetentionPolicy, Get-RetentionPolicyTag,',
        'Get-IRMConfiguration, Get-OMEConfiguration, Get-InboxRule',
        '',
        'Purview (Security & Compliance, -IncludePurview): Get-ProtectionAlert, Get-RetentionCompliancePolicy/Rule,',
        'Get-DlpCompliancePolicy/Rule, Get-Label, Get-LabelPolicy, Get-ComplianceTag',
        '',
        'Graph (-IncludeGraph): Get-MgSecuritySecureScore, Get-MgDirectoryRole, Get-MgDirectoryRoleMember'
    )
    foreach ($line in $cmdletList) { Add-Line $line }
    Add-Line
}

# ============================================================
# Main
# ============================================================

$safeCustomer = ($CustomerName -replace '[\\/:*?"<>|\s]', '_')
$dateStamp = (Get-Date).ToString('yyyyMMdd')
$reportName = "EXO-Review_${safeCustomer}_${dateStamp}.md"
$resolvedOutput = Resolve-Path -LiteralPath $OutputPath -ErrorAction SilentlyContinue
if (-not $resolvedOutput) {
    New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
    $resolvedOutput = Resolve-Path -LiteralPath $OutputPath
}
$reportPath = Join-Path $resolvedOutput.Path $reportName
$csvDir = Join-Path $resolvedOutput.Path "EXO-Review_${safeCustomer}_${dateStamp}_csv"

if ($AppId -and (-not $CertificateThumbprint -or -not $Organization)) {
    Write-Error "-AppId requires -CertificateThumbprint and -Organization."
    exit 1
}

try {
    Initialize-ExchangeOnlineModule

    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($UserPrincipalName) { $exoParams.UserPrincipalName = $UserPrincipalName }
    if ($AppId) {
        $exoParams.AppId = $AppId
        $exoParams.CertificateThumbprint = $CertificateThumbprint
        $exoParams.Organization = $Organization
    }
    try {
        Connect-ExchangeOnline @exoParams
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }

    # Pre-collect tenant identity for the header
    try {
        $script:OrgConfig = Get-OrganizationConfig -ErrorAction Stop
        $script:TenantName = $script:OrgConfig.DisplayName
    }
    catch { Add-LogEntry -Section 'OrganizationConfig (header)' -Reason $_.Exception.Message }
    try {
        $script:AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop)
        $script:InitialDomain = @($script:AcceptedDomains | Where-Object { $_.InitialDomain } | Select-Object -First 1).DomainName
    }
    catch { Add-LogEntry -Section 'AcceptedDomain (header)' -Reason $_.Exception.Message }

    # Collected-by identity
    try {
        $connInfo = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
        if ($connInfo.UserPrincipalName) { $script:CollectedBy = $connInfo.UserPrincipalName }
        elseif ($connInfo.User) { $script:CollectedBy = $connInfo.User }
    }
    catch { }

    if ($IncludePurview) {
        Write-Host "Connecting to Security & Compliance PowerShell (Purview)..." -ForegroundColor Cyan
        $ippsParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
        if ($UserPrincipalName) { $ippsParams.UserPrincipalName = $UserPrincipalName }
        if ($AppId) {
            $ippsParams.AppId = $AppId
            $ippsParams.CertificateThumbprint = $CertificateThumbprint
            $ippsParams.Organization = $Organization
        }
        try {
            Connect-IPPSSession @ippsParams
            $script:PurviewConnected = $true
            Write-Host "Connected to Security & Compliance." -ForegroundColor Green
        }
        catch {
            $script:PurviewError = $_.Exception.Message
            Write-Warning "Could not connect to Security & Compliance PowerShell - $script:PurviewError. Purview sections will be marked Not available."
            Add-LogEntry -Section 'Purview connection' -Reason $script:PurviewError
        }
    }

    if ($IncludeGraph) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
            try {
                Connect-MgGraph -Scopes 'SecurityEvents.Read.All','RoleManagement.Read.Directory','Directory.Read.All' -NoWelcome -ErrorAction Stop
                $script:GraphConnected = $true
                Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
            }
            catch {
                $script:GraphError = $_.Exception.Message
                Write-Warning "Could not connect to Microsoft Graph - $script:GraphError."
                Add-LogEntry -Section 'Graph connection' -Reason $script:GraphError
            }
        }
        else {
            $script:GraphError = 'Microsoft.Graph modules not installed'
            Write-Warning "Microsoft.Graph modules not found. Graph sections will be marked Not available."
            Add-LogEntry -Section 'Graph connection' -Reason $script:GraphError
        }
    }

    # Header (section 0)
    $purviewState = if (-not $IncludePurview) { 'No' } elseif ($script:PurviewConnected) { 'Yes' } else { "Failed ($script:PurviewError)" }
    $graphState = if (-not $IncludeGraph) { 'No' } elseif ($script:GraphConnected) { 'Yes' } else { "Failed ($script:GraphError)" }

    Add-Line "# Exchange Online Review - $CustomerName"
    Add-Line
    Add-Line "| Setting | Value |"
    Add-Line "|---|---|"
    Add-Line "| Customer | $(Convert-ValueToText $CustomerName) |"
    Add-Line "| Tenant display name | $(Convert-ValueToText $script:TenantName) |"
    Add-Line "| Initial domain | $(Convert-ValueToText $script:InitialDomain) |"
    Add-Line "| Review date | $($ReviewDate.ToString('yyyy-MM-dd')) |"
    Add-Line "| Reviewed by | $(Convert-ValueToText $ReviewedBy) |"
    Add-Line "| Collected by | $(Convert-ValueToText $script:CollectedBy) |"
    Add-Line "| EXO module version | $(Convert-ValueToText $script:ExoModuleVersion) |"
    Add-Line "| Script version | $script:ScriptVersion |"
    Add-Line "| Purview data collected | $purviewState |"
    Add-Line "| Graph data collected | $graphState |"
    Add-Line "| Mailbox statistics collected | $(if ($IncludeMailboxStatistics) { 'Yes' } else { 'No' }) |"
    Add-Line "| Message trace collected | $(if ($IncludeMessageTrace) { 'Yes' } else { 'No' }) |"
    Add-Line
    Add-Line "## Table of contents"
    Add-Line
    foreach ($title in $script:SectionTitles.Values) {
        Add-Line "- [$title](#$(ConvertTo-MdAnchor $title))"
    }
    Add-Line
    $summaryInsertAt = $script:Report.Length

    Write-Host "`nCollecting report data..." -ForegroundColor Cyan
    Get-OrgConfigSection
    Get-RecipientsSection
    Get-MailboxHygieneSection
    Get-GroupsSection
    Get-DomainsSection
    Get-MailFlowSection
    Get-HybridSection
    Get-SharingSection
    Get-ClientAccessSection
    Get-PermissionsSection
    Get-ThreatProtectionSection
    Get-AlertPoliciesSection
    Get-ComplianceSection
    Get-GraphSection

    Write-Appendix

    # Insert the summary counts block right after the TOC (header position)
    $summaryBlock = [System.Text.StringBuilder]::new()
    [void]$summaryBlock.AppendLine("## $($script:SectionTitles.Summary)")
    [void]$summaryBlock.AppendLine()
    if ($script:Summary.Count -eq 0) {
        [void]$summaryBlock.AppendLine("_No counts collected._")
    }
    else {
        $sumRows = @($script:Summary.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{ Metric = $_.Key; Count = $_.Value }
        })
        [void]$summaryBlock.AppendLine((ConvertTo-MdTable -Rows $sumRows -Columns @('Metric','Count')))
    }
    [void]$summaryBlock.AppendLine()
    [void]$script:Report.Insert($summaryInsertAt, $summaryBlock.ToString())

    # Write report
    $script:Report.ToString() | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "`nReport written to: $reportPath" -ForegroundColor Green

    # CSV export
    if ($ExportCsv -and $script:CsvData.Count -gt 0) {
        New-Item -ItemType Directory -Force -Path $csvDir | Out-Null
        foreach ($key in $script:CsvData.Keys) {
            $csvPath = Join-Path $csvDir "$key.csv"
            $script:CsvData[$key] | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "CSV exports written to: $csvDir" -ForegroundColor Green
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    if ($script:GraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue }
    Write-Host "Done." -ForegroundColor Green
}
