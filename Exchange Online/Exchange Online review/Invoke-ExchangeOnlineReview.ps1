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

.PARAMETER ExoLogPath
    Optional folder for ExchangeOnlineManagement client logs. When set,
    Connect-ExchangeOnline and Connect-IPPSSession run with
    -EnableErrorReporting -LogDirectoryPath <path> -LogLevel All. Useful
    when cmdlets fail with the generic 'server side error' message.

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
    Version:    1.2
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
    [string]$DnsServer,

    [Parameter(Mandatory = $false)]
    [string]$ExoLogPath
)

$script:ScriptVersion = "1.2"
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
        Write-Error "ExchangeOnlineManagement module not found. Install it first: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
        exit 1
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
    param([string]$Section, [string]$Reason, [System.Management.Automation.ErrorRecord]$ErrorRecord)
    if ($ErrorRecord) { $Reason = Get-ErrorDetail $ErrorRecord }
    $script:CollectionLog.Add([PSCustomObject]@{
        Section = $Section
        Reason  = $Reason
        Time    = (Get-Date).ToString("HH:mm:ss")
    })
    Write-Host "  [$Section] $Reason" -ForegroundColor DarkYellow
}

function Get-ErrorDetail {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $messages = @(); $ex = $ErrorRecord.Exception
    while ($ex) { $messages += $ex.Message; $ex = $ex.InnerException }
    $text = ($messages | Select-Object -Unique) -join ' --> '
    if ($ErrorRecord.FullyQualifiedErrorId) { $text += " [$($ErrorRecord.FullyQualifiedErrorId)]" }
    return $text
}

function Invoke-ExoProxy {
    $pass = @()
    for ($i = 1; $i -lt $args.Count; $i++) {
        if ($args[$i] -is [string] -and $args[$i] -match '^-ErrorAction:?$') { $i++; continue }
        $pass += , $args[$i]
    }
    $PSDefaultParameterValues = @{}
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $sp = [Net.ServicePointManager]::SecurityProtocol
        if (-not ($sp -band [Net.SecurityProtocolType]::Tls12)) { [Net.ServicePointManager]::SecurityProtocol = $sp -bor [Net.SecurityProtocolType]::Tls12 }
        $ev = $null
        $out = & $script:ExoProxyCommands[$args[0]] @pass -ErrorVariable ev 2>$null
        $records = @($ev | ForEach-Object {
            if ($_ -is [System.Management.Automation.ErrorRecord]) { $_ }
            elseif ($_ -is [System.Management.Automation.IContainsErrorRecord]) { $_.ErrorRecord }
            elseif ($_ -is [Exception]) { [System.Management.Automation.ErrorRecord]::new($_, 'ExoProxyError', 'NotSpecified', $null) }
        })
        $tlsBroken = @($records | Where-Object { "$($_.Exception)" -match 'SslProtocolType' }).Count -gt 0
        if (-not ($tlsBroken -and $attempt -eq 1)) { break }
    }
    $final = @($records | Where-Object {
        $x = $_.Exception; $benign = $false
        while ($x) {
            if ($x -is [System.Net.WebException]) { $benign = ($x.Response -and [int]$x.Response.StatusCode -eq 403); break }
            $x = $x.InnerException
        }
        -not $benign
    })
    if ($final.Count -gt 0) {
        throw ((@($records) | ForEach-Object { Get-ErrorDetail $_ } | Select-Object -Unique) -join ' | ')
    }
    $out
}

function Register-ExoProxyWrapper {
    $script:ExoProxyCommands = @{}
    $names = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ModuleName } | Where-Object { $_ })
    $candidates = @($names) + @($names | ForEach-Object { Split-Path $_ -Leaf }) + @($names | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) })
    $modules = @(Get-Module | Where-Object { $candidates -contains $_.Name -or $candidates -contains $_.Path -or $candidates -contains $_.ModuleBase -or $_.Name -like 'tmpEXO_*' } | Sort-Object Name -Unique)
    foreach ($mod in $modules) {
        foreach ($cmd in @(Get-Command -Module $mod.Name -CommandType Function)) {
            $script:ExoProxyCommands[$cmd.Name] = "$($mod.Name)\$($cmd.Name)"
            Set-Item -Path "function:script:$($cmd.Name)" -Value ([scriptblock]::Create("Invoke-ExoProxy '$($cmd.Name)' @args"))
        }
    }
    if ($script:ExoProxyCommands.Count -eq 0) { Write-Warning "No Exchange Online proxy cmdlets found to wrap; cmdlets may fail with -ErrorAction Stop." }
    else { Write-Host "Prepared $($script:ExoProxyCommands.Count) Exchange Online / Purview cmdlets." -ForegroundColor DarkGray }
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

function Resolve-DirectoryObjectLabel {
    param([string]$Id)
    if ($null -eq $script:DirObjCache) { $script:DirObjCache = @{} }
    if ($script:DirObjCache.ContainsKey($Id)) { return $script:DirObjCache[$Id] }
    if ($null -eq $script:SpCache) { $script:SpCache = @(); try { $script:SpCache = @(Get-ServicePrincipal) } catch { } }
    $label = $null
    $sp = @($script:SpCache | Where-Object { "$($_.ObjectId)" -eq $Id -or "$($_.AppId)" -eq $Id -or "$($_.Identity)" -eq $Id }) | Select-Object -First 1
    if ($sp) { $label = "$($sp.DisplayName) (app)" }
    if (-not $label) { try { $u = Get-User -Identity $Id -ErrorAction Stop; if ($u) { $label = "$($u.DisplayName)" } } catch { } }
    if (-not $label) { try { $g = Get-Group -Identity $Id -ErrorAction Stop; if ($g) { $label = "$($g.DisplayName) (group)" } } catch { } }
    $namePart = "$($label -replace ' \((app|group)\)$', '')".Trim()
    if (-not $label -or $namePart -notmatch '\S' -or $namePart -match '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') { $label = "Unresolved Entra object $Id" }
    $script:DirObjCache[$Id] = $label
    return $label
}

function Get-RoleGroupDisplayName {
    param([string]$Name)
    $map = [ordered]@{ '^TenantAdmins_' = 'Global Administrator'; '^ExchangeServiceAdmins_' = 'Exchange Administrator'; '^ComplianceAdmins_' = 'Compliance Administrator'; '^SecurityAdmins_' = 'Security Administrator'; '^GlobalReaders_' = 'Global Reader' }
    foreach ($k in $map.Keys) { if ($Name -match $k) { return "$Name (Entra role: $($map[$k]))" } }
    if ($Name -match '^[A-Za-z]+_-?\d{6,}$') { return "$Name (linked to an Entra role)" }
    return $Name
}

function ConvertTo-CsvRow {
    param([object[]]$Rows)
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $o = [ordered]@{}
        foreach ($p in $row.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) {
                $v = (@($v | ForEach-Object { if ($null -ne $_.Name -and $_ -isnot [string]) { "$($_.Name)" } else { "$_" } }) -join '; ')
            }
            $o[$p.Name] = $v
        }
        [PSCustomObject]$o
    }
}
function Test-IsTrue {
    param($Value)
    if ($Value -is [bool]) { return $Value }
    return ("$Value" -eq 'True')
}

function Format-RetentionPeriod {
    param($Tag)
    if (-not (Test-IsTrue $Tag.RetentionEnabled) -or $null -eq $Tag.AgeLimitForRetention -or "$($Tag.AgeLimitForRetention)" -eq '') { return 'Unlimited' }
    [timespan]$ts = [timespan]::Zero
    if ($Tag.AgeLimitForRetention -is [timespan]) { $ts = $Tag.AgeLimitForRetention }
    elseif (-not [timespan]::TryParse("$($Tag.AgeLimitForRetention)", [ref]$ts)) { return "$($Tag.AgeLimitForRetention)" }
    $d = [int]$ts.TotalDays
    if ($d -gt 0 -and $d % 365 -eq 0) { return "$d days ($($d / 365) year(s))" }
    return "$d days"
}
function Format-RetentionRule {
    param($Rule, [hashtable]$TagNames)
    $label = "$(@($Rule.PublishComplianceTag, $Rule.ApplyComplianceTag) | Where-Object { $_ } | Select-Object -First 1)"
    $label = ($label -split ',')[0].Trim()
    if ($TagNames -and $TagNames.ContainsKey($label)) { $label = $TagNames[$label] }
    if ($Rule.PublishComplianceTag) { return "Publishes retention label: $label" }
    if ($Rule.ApplyComplianceTag) { return "Auto-applies retention label: $label" }
    $dur = "$($Rule.RetentionDuration)"
    $d = 0
    $durText = if ($dur -eq '' -or $dur -eq 'Unlimited') { 'forever' }
               elseif ([int]::TryParse($dur, [ref]$d)) { if ($d -gt 0 -and $d % 365 -eq 0) { "$d days ($($d / 365) year(s))" } else { "$d days" } }
               else { $dur }
    $basis = switch ("$($Rule.ExpirationDateOption)") {
        'CreationAgeInDays'     { ' (based on when items were created)' }
        'ModificationAgeInDays' { ' (based on when items were last modified)' }
        default                 { '' }
    }
    switch ("$($Rule.RetentionComplianceAction)") {
        'Keep'          { return $(if ($durText -eq 'forever') { "Retain items forever$basis" } else { "Retain items for $durText$basis" }) }
        'Delete'        { return "Delete items older than $durText$basis" }
        'KeepAndDelete' { return "Retain items for $durText, then delete them$basis" }
        default         { return "$($Rule.Name)" }
    }
}
function Get-SensitivityLabelInfo {
    param($Label, [object[]]$AllLabels)
    $types = @(@($Label.LabelActions) | ForEach-Object { if ("$_" -match '"Type"\s*:\s*"([^"]+)"') { $Matches[1].ToLower() } })
    $scopeMap = @{ 'File' = 'Files & other data assets'; 'Email' = 'Email'; 'Site' = 'Site'; 'UnifiedGroup' = 'UnifiedGroup'; 'Teamwork' = 'Meetings'; 'SchematizedData' = 'Schematized data assets' }
    $scope = @("$($Label.ContentType)" -split '\s*,\s*' | Where-Object { $_ } | ForEach-Object { if ($scopeMap.ContainsKey($_)) { $scopeMap[$_] } else { $_ } }) -join ', '
    $findLabel = { param($id) @($AllLabels | Where-Object { "$id" -ne '' -and ("$($_.Guid)" -eq "$id" -or "$($_.ImmutableId)" -eq "$id" -or "$($_.Name)" -eq "$id" -or "$($_.Identity)" -eq "$id") }) | Select-Object -First 1 }
    $parent = $null
    if ("$($Label.ParentId)" -match '\S') { $p = & $findLabel $Label.ParentId; $parent = if ($p) { "$($p.DisplayName)" } else { "$($Label.ParentId)" } }
    $sub = @($AllLabels | Where-Object { "$($_.ParentId)" -match '\S' -and ((& $findLabel $_.ParentId) -eq $Label) } | Sort-Object Priority | ForEach-Object { "$($_.DisplayName)" })
    $marking = @()
    if ($types -contains 'applycontentmarking' -or $types -contains 'applycontentmarkingheader' -or $Label.ApplyContentMarkingHeaderEnabled) { $marking += 'Header' }
    if ($types -contains 'applycontentmarkingfooter' -or $Label.ApplyContentMarkingFooterEnabled) { $marking += 'Footer' }
    if ($types -contains 'applywatermarking' -or $Label.ApplyWaterMarkingEnabled) { $marking += 'Watermark' }
    $encrypt = ($types -contains 'encrypt') -or $Label.EncryptionEnabled
    $groupSite = [bool]$Label.SiteAndGroupProtectionEnabled
    [PSCustomObject]@{
        Scope          = $scope
        Parent         = $parent
        Sublabels      = $sub
        AccessControl  = $(if ($encrypt) { 'Access control (encryption)' } else { 'None' })
        ContentMarking = $(if ($marking.Count) { $marking -join ', ' } else { 'None' })
        AutoLabeling   = $(if ("$($Label.Conditions)" -match '\S') { 'Configured' } else { 'None' })
        GroupSettings  = $(if (($types -contains 'protectgroup') -or ($groupSite -and $scope -match 'UnifiedGroup')) { 'Configured' } else { 'None' })
        SiteSettings   = $(if (($types -contains 'protectsite') -or ($groupSite -and $scope -match 'Site')) { 'Configured' } else { 'None' })
        ActionTypes    = ($types | Select-Object -Unique) -join ', '
    }
}

function Resolve-LabelNames {
    param([object[]]$Ids, [object[]]$AllLabels)
    @($Ids | Where-Object { $_ } | ForEach-Object {
        $id = "$_"
        $l = @($AllLabels | Where-Object { "$($_.Guid)" -eq $id -or "$($_.ImmutableId)" -eq $id -or "$($_.Name)" -eq $id -or "$($_.Identity)" -eq $id }) | Select-Object -First 1
        if ($l) { "$($l.DisplayName)" } else { $id }
    })
}
function Get-RoleGroupMemberInfo {
    param([object[]]$Members)
    $groups = @(); $users = @()
    $guid = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    foreach ($m in @($Members | Where-Object { $_ })) {
        $label = @("$($m.DisplayName)", "$($m.PrimarySmtpAddress)", "$($m.Name)") | Where-Object { $_ -and $_ -notmatch $guid } | Select-Object -First 1
        if (-not $label) { $label = Resolve-DirectoryObjectLabel -Id "$(@($m.ExternalDirectoryObjectId, $m.Name, $m.Identity) | Where-Object { "$_" -match $guid } | Select-Object -First 1)" }
        if ("$($m.RecipientType) $($m.RecipientTypeDetails)" -match 'Group' -or $label -like '* (group)') { $groups += $label } else { $users += $label }
    }
    $parts = @()
    if ($groups.Count) { $parts += "$($groups.Count) group(s): " + (@($groups | Select-Object -First 3) -join ', ') + $(if ($groups.Count -gt 3) { ' …' } else { '' }) }
    if ($users.Count) { $parts += "$($users.Count) user(s)/app(s)" }
    [PSCustomObject]@{
        Summary = $(if ($parts.Count) { $parts -join '; ' } else { 'No members' })
        Groups  = $groups
        Users   = $users
    }
}
function Get-QuarantinePermissionInfo {
    param($Policy)
    $map = [ordered]@{
        'PermissionToRelease'        = @{ Bit = 4;  Label = 'Release the message from quarantine' }
        'PermissionToRequestRelease' = @{ Bit = 8;  Label = 'Request release of the message' }
        'PermissionToDelete'         = @{ Bit = 1;  Label = 'Delete the message' }
        'PermissionToPreview'        = @{ Bit = 2;  Label = 'Preview the message' }
        'PermissionToAllowSender'    = @{ Bit = 32; Label = 'Allow sender' }
        'PermissionToBlockSender'    = @{ Bit = 16; Label = 'Block sender' }
    }
    $v = 0
    $text = "$($Policy.EndUserQuarantinePermissions)"
    if ($text -match 'PermissionTo\w+\s*:') {
        foreach ($k in $map.Keys) { if ($text -match "$k\s*:\s*True") { $v = $v -bor $map[$k].Bit } }
    }
    else { [void][int]::TryParse("$($Policy.EndUserQuarantinePermissionsValue)", [ref]$v) }
    $granted = @($map.Keys | Where-Object { $v -band $map[$_].Bit } | ForEach-Object { $map[$_].Label })
    $access = switch ($v -band 63) {
        0       { 'No access' }
        43      { 'Limited access' }
        39      { 'Full access' }
        default { 'Specific access (custom)' }
    }
    [PSCustomObject]@{
        Access        = $access
        Permissions   = $(if ($granted.Count) { $granted -join '; ' } else { 'None (view message header only)' })
        Notifications = $(if (Test-IsTrue $Policy.ESNEnabled) { 'Enabled' } else { 'Disabled' })
    }
}
function Get-PolicyRuleInfo {
    param($Policy, $Rule, [object[]]$PresetRules, [object[]]$BuiltInRules, [switch]$SenderBased, [switch]$RuleLookupFailed)
    $f = if ($SenderBased) { @('From','FromMemberOf','SenderDomainIs','ExceptIfFrom','ExceptIfFromMemberOf','ExceptIfSenderDomainIs') }
         else { @('SentTo','SentToMemberOf','RecipientDomainIs','ExceptIfSentTo','ExceptIfSentToMemberOf','ExceptIfRecipientDomainIs') }
    $refProps = 'HostedContentFilterPolicy','AntiPhishPolicy','MalwareFilterPolicy','SafeAttachmentPolicy','SafeLinksPolicy'
    $matchRule = { param($rules) @($rules | Where-Object { $r = $_; @($refProps | Where-Object { "$($r.$_)" -eq "$($Policy.Name)" }).Count -gt 0 }) | Select-Object -First 1 }
    $kind = 'Custom'; $scopeRule = $Rule
    if ("$($Policy.Name)" -eq 'Evaluation Policy') {
        $kind = 'Evaluation'
    }
    elseif ("$($Policy.RecommendedPolicyType)" -in 'Standard','Strict') {
        $kind = "$($Policy.RecommendedPolicyType) preset"; $scopeRule = & $matchRule $PresetRules
    }
    elseif ((Test-IsTrue $Policy.IsBuiltInProtection) -or "$($Policy.Name)" -like 'Built-In Protection Policy*') {
        $kind = 'Built-in protection'; $scopeRule = & $matchRule $BuiltInRules
    }
    elseif ((Test-IsTrue $Policy.IsDefault) -or "$($Policy.Name)" -eq 'Default') { $kind = 'Default' }
    $status = switch ($kind) {
        'Default'             { 'On (default policy)' }
        'Built-in protection' { 'On (built-in protection)' }
        'Evaluation'          { 'Not in use (evaluation policy)' }
        default { if ($scopeRule) { "$($scopeRule.State)" } elseif ($RuleLookupFailed -and -not $scopeRule) { 'Unknown (rule lookup failed)' } else { 'Not applied (no rule)' } }
    }
    $inc = @(@($scopeRule.($f[0])) + @($scopeRule.($f[1])) + @($scopeRule.($f[2])) | Where-Object { $_ })
    $exc = @(@($scopeRule.($f[3])) + @($scopeRule.($f[4])) + @($scopeRule.($f[5])) | Where-Object { $_ })
    $applies = if ($kind -eq 'Evaluation') { 'Not applicable' }
        elseif ($kind -in 'Default','Built-in protection' -and $inc.Count -eq 0) { 'All recipients' }
        elseif ($kind -in 'Standard preset','Strict preset' -and $scopeRule -and $inc.Count -eq 0) { 'All recipients' }
        elseif ($RuleLookupFailed -and -not $scopeRule) { 'Unknown (rule lookup failed)' }
        elseif ($inc.Count -eq 0) { 'Nobody (no conditions)' }
        else {
            $parts = @()
            $u = @($scopeRule.($f[0]) | Where-Object { $_ }).Count; if ($u) { $parts += "$u user(s)" }
            $g = @($scopeRule.($f[1]) | Where-Object { $_ }).Count; if ($g) { $parts += "$g group(s)" }
            $d = @($scopeRule.($f[2]) | Where-Object { $_ }); if ($d.Count) { $parts += "Domains: " + ((@($d | Select-Object -First 3)) -join ', ') + $(if ($d.Count -gt 3) { " (+$($d.Count - 3))" } else { '' }) }
            $parts -join '; '
        }
    if ($exc.Count -gt 0) { $applies += ' (with exclusions)' }
    [PSCustomObject]@{
        Kind = $kind; Status = $status; Priority = $(if ($kind -in 'Default','Built-in protection') { 'Lowest' } elseif ($kind -eq 'Evaluation') { $null } elseif ($scopeRule) { $scopeRule.Priority } else { $null }); AppliesTo = $applies
        IncludedUsers = $scopeRule.($f[0]); IncludedGroups = $scopeRule.($f[1]); IncludedDomains = $scopeRule.($f[2])
        ExcludedUsers = $scopeRule.($f[3]); ExcludedGroups = $scopeRule.($f[4]); ExcludedDomains = $scopeRule.($f[5])
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
    $PSDefaultParameterValues = @{ '*:ErrorAction' = 'Stop' }
    try {
        $result = & $Body
        if ($result) {
            if ($result -is [string]) { Add-Line $result } else { Add-Line ($result | Out-String) }
        }
    }
    catch {
        $reason = Get-ErrorDetail $_
        Add-Line "_Not available: $($reason)_"
        Add-Line
        Add-LogEntry -Section $Title -Reason $reason
    }
    Add-Line
}

function Resolve-DnsSafe {
    param([string]$Name, [string]$Type)
    if ($null -eq $script:DnsBadServers) { $script:DnsBadServers = @{} }
    $servers = if ($DnsServer) { @($DnsServer) } else { @('1.1.1.1', '8.8.8.8', '') }
    $lastError = $null
    foreach ($server in $servers) {
        if ($server -and $script:DnsBadServers.ContainsKey($server)) { continue }
        $params = @{ Name = $Name; Type = $Type; DnsOnly = $true; QuickTimeout = $true; ErrorAction = "Stop" }
        if ($server) { $params.Server = $server }
        try {
            return @(DnsClient\Resolve-DnsName @params | Where-Object { $_.Section -eq 'Answer' -and "$($_.Type)" -eq $Type })
        }
        catch {
            # NXDOMAIN is an expected "no record" result; anything else is logged so empty DNS cells can be explained.
            if ($_.Exception.Message -match 'does not exist') { return @() }
            $lastError = $_
            if ($server -and -not $DnsServer) { $script:DnsBadServers[$server] = $true }
        }
    }
    if ($lastError) { Add-LogEntry -Section "DNS $Type $Name" -ErrorRecord $lastError }
    return @()
}

function Get-TxtRecords {
    param([string]$Name)
    $records = Resolve-DnsSafe -Name $Name -Type TXT
    return @($records | ForEach-Object { ($_.Strings -join "") })
}

function Get-MessageDirection {
    # Classifies a message against the tenant's accepted domains (lowercase list).
    param([string]$SenderAddress, [string]$RecipientAddress, [string[]]$Domains)
    $s = ("$SenderAddress" -split '@')[-1].ToLower()
    $r = ("$RecipientAddress" -split '@')[-1].ToLower()
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
        $props = @('PopEnabled', 'ImapEnabled', 'EwsEnabled', 'ActiveSyncEnabled', 'MAPIEnabled', 'OWAEnabled',
                   'SmtpClientAuthenticationDisabled', 'OwaMailboxPolicy', 'ActiveSyncMailboxPolicy')
        $script:CasMailboxes = @(Get-EXOCASMailbox -ResultSize Unlimited -Properties $props)
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
        try { $inactive = @(Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Inactive mailboxes' -ErrorRecord $_ }
        try { $softDeleted = @(Get-EXOMailbox -SoftDeletedMailbox -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Soft-deleted mailboxes' -ErrorRecord $_ }
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
            [PSCustomObject]@{ State = 'Litigation hold enabled'; Count = @($mbx | Where-Object { Test-IsTrue $_.LitigationHoldEnabled }).Count }
            [PSCustomObject]@{ State = 'Archive enabled'; Count = @($mbx | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count }
            [PSCustomObject]@{ State = 'Auto-expanding archive'; Count = @($mbx | Where-Object { Test-IsTrue $_.AutoExpandingArchiveEnabled }).Count }
            [PSCustomObject]@{ State = 'Retention hold enabled'; Count = @($mbx | Where-Object { Test-IsTrue $_.RetentionHoldEnabled }).Count }
        )

        Add-Line "### Mailbox auditing"
        Add-Line
        Out-Table -CsvName "MailboxAuditCounts" -Columns @('State','Count') -Rows @(
            [PSCustomObject]@{ State = 'Audit enabled'; Count = @($mbx | Where-Object { Test-IsTrue $_.AuditEnabled }).Count }
            [PSCustomObject]@{ State = 'Audit disabled'; Count = @($mbx | Where-Object { -not (Test-IsTrue $_.AuditEnabled) }).Count }
        )
        $bypass = @()
        try {
            Assert-Cmdlet Get-MailboxAuditBypassAssociation
            $bypass = @(Get-MailboxAuditBypassAssociation -ResultSize Unlimited | Where-Object { Test-IsTrue $_.AuditBypassEnabled })
        }
        catch { Add-LogEntry -Section 'Mailbox audit bypass' -ErrorRecord $_ }
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
        catch { Add-LogEntry -Section 'Shared/room sign-in state' -ErrorRecord $_ }
        $enabledRows = @($users | Where-Object { -not (Test-IsTrue $_.AccountDisabled) } | ForEach-Object {
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
        try { $dyn = @(Get-DynamicDistributionGroup -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Dynamic distribution groups' -ErrorRecord $_ }
        $m365 = @()
        try { $m365 = @(Get-UnifiedGroup -ResultSize Unlimited) } catch { Add-LogEntry -Section 'Microsoft 365 groups' -ErrorRecord $_ }

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
                AllowExternalSenders = (-not (Test-IsTrue $_.RequireSenderAuthenticationEnabled))
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
        $extRows = @($dg | Where-Object { -not (Test-IsTrue $_.RequireSenderAuthenticationEnabled) } | ForEach-Object {
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
        try { $dkim = @(Get-DkimSigningConfig) } catch { Add-LogEntry -Section 'DKIM config' -ErrorRecord $_ }

        Add-Line "### DNS records per domain"
        Add-Line
        $dnsRows = @()
        foreach ($d in $domains) {
            $name = "$($d.DomainName)".Trim()
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
        try { $arc = @(Get-ArcConfig) } catch { Add-LogEntry -Section 'ARC config' -ErrorRecord $_ }
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
        try { $inbound = @(Get-InboundConnector) } catch { Add-LogEntry -Section 'Inbound connectors' -ErrorRecord $_ }
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
        try { $outbound = @(Get-OutboundConnector) } catch { Add-LogEntry -Section 'Outbound connectors' -ErrorRecord $_ }
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
        try { $rules = @(Get-TransportRule) } catch { Add-LogEntry -Section 'Transport rules' -ErrorRecord $_ }
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
        try { $remote = @(Get-RemoteDomain) } catch { Add-LogEntry -Section 'Remote domains' -ErrorRecord $_ }
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
        try { $journal = @(Get-JournalRule) } catch { Add-LogEntry -Section 'Journal rules' -ErrorRecord $_ }
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
        try { $hve = @(Get-MailUser -HVEAccount -ResultSize Unlimited) } catch { Add-LogEntry -Section 'HVE accounts' -ErrorRecord $_ }
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
            $dirCounts[(Get-MessageDirection -SenderAddress $msg.SenderAddress -RecipientAddress $msg.RecipientAddress -Domains $acceptedNames)]++
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
        try { $opo = @(Get-OnPremisesOrganization) } catch { $opo = @(); Add-LogEntry -Section 'OnPremisesOrganization' -ErrorRecord $_ }
        $rows += @($opo | ForEach-Object {
            [PSCustomObject]@{ Type = 'OnPremisesOrganization'; Name = $_.Name; Details = "HybridDomains=$($_.HybridDomains); Guid=$($_.Guid)" }
        })
        try { $ioc = @(Get-IntraOrganizationConnector) } catch { $ioc = @(); Add-LogEntry -Section 'IntraOrganizationConnector' -ErrorRecord $_ }
        $rows += @($ioc | ForEach-Object {
            [PSCustomObject]@{ Type = 'IntraOrganizationConnector'; Name = $_.Name; Details = "Enabled=$($_.Enabled); TargetAddressDomains=$($_.TargetAddressDomains); DiscoveryEndpoint=$($_.DiscoveryEndpoint)" }
        })
        try { $mep = @(Get-MigrationEndpoint) } catch { $mep = @(); Add-LogEntry -Section 'MigrationEndpoint' -ErrorRecord $_ }
        $rows += @($mep | ForEach-Object {
            [PSCustomObject]@{ Type = 'MigrationEndpoint'; Name = $_.Identity; Details = "RemoteServer=$($_.RemoteServer); ExchangeVersion=$($_.ExchangeVersion)" }
        })
        try { $batch = @(Get-MigrationBatch) } catch { $batch = @(); Add-LogEntry -Section 'MigrationBatch' -ErrorRecord $_ }
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
        try { $rel = @(Get-OrganizationRelationship) } catch { Add-LogEntry -Section 'OrganizationRelationship' -ErrorRecord $_ }
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
        try { $sp = @(Get-SharingPolicy) } catch { Add-LogEntry -Section 'SharingPolicy' -ErrorRecord $_ }
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
        try { $aas = @(Get-AvailabilityAddressSpace) } catch { Add-LogEntry -Section 'AvailabilityAddressSpace' -ErrorRecord $_ }
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
        try { $owa = @(Get-OwaMailboxPolicy) } catch { Add-LogEntry -Section 'OwaMailboxPolicy' -ErrorRecord $_ }
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
        try { $aso = @(Get-ActiveSyncOrganizationSettings) } catch { Add-LogEntry -Section 'ActiveSyncOrganizationSettings' -ErrorRecord $_ }
        $asoRows = @($aso | ForEach-Object {
            [PSCustomObject]@{ DefaultAccessLevel = $_.DefaultAccessLevel; UserMailInsert = $_.UserMailInsert; AdminMailRecipients = $_.AdminMailRecipients }
        })
        Out-Table -CsvName "ActiveSyncOrgSettings" -Rows $asoRows -Columns @('DefaultAccessLevel','UserMailInsert','AdminMailRecipients')
        $asr = @()
        try { $asr = @(Get-ActiveSyncDeviceAccessRule) } catch { Add-LogEntry -Section 'ActiveSyncDeviceAccessRule' -ErrorRecord $_ }
        $asrRows = @($asr | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Characteristic = $_.Characteristic; QueryString = $_.QueryString; AccessLevel = $_.AccessLevel }
        })
        Out-Table -CsvName "ActiveSyncDeviceAccessRules" -Rows $asrRows -Columns @('Name','Characteristic','QueryString','AccessLevel')

        Add-Line "### Mobile device mailbox policies"
        Add-Line
        $mdp = @()
        try { $mdp = @(Get-MobileDeviceMailboxPolicy) } catch { Add-LogEntry -Section 'MobileDeviceMailboxPolicy' -ErrorRecord $_ }
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
        try { $cas = Get-SharedCasMailboxes } catch { Add-LogEntry -Section 'CAS mailboxes' -ErrorRecord $_ }
        Out-Table -CsvName "ProtocolUsage" -Columns @('Protocol','Enabled','Disabled','FollowsOrgSetting') -Rows @(
            [PSCustomObject]@{ Protocol = 'POP'; Enabled = @($cas | Where-Object { Test-IsTrue $_.PopEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.PopEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'IMAP'; Enabled = @($cas | Where-Object { Test-IsTrue $_.ImapEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.ImapEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'EWS'; Enabled = @($cas | Where-Object { Test-IsTrue $_.EwsEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.EwsEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'ActiveSync'; Enabled = @($cas | Where-Object { Test-IsTrue $_.ActiveSyncEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.ActiveSyncEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'MAPI'; Enabled = @($cas | Where-Object { Test-IsTrue $_.MapiEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.MapiEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'OWA'; Enabled = @($cas | Where-Object { Test-IsTrue $_.OWAEnabled }).Count; Disabled = @($cas | Where-Object { -not (Test-IsTrue $_.OWAEnabled) }).Count; FollowsOrgSetting = $null }
            [PSCustomObject]@{ Protocol = 'SMTP AUTH'; Enabled = @($cas | Where-Object { "$($_.SmtpClientAuthenticationDisabled)" -eq 'False' }).Count; Disabled = @($cas | Where-Object { "$($_.SmtpClientAuthenticationDisabled)" -eq 'True' }).Count; FollowsOrgSetting = @($cas | Where-Object { $null -eq $_.SmtpClientAuthenticationDisabled -or "$($_.SmtpClientAuthenticationDisabled)" -eq '' }).Count }
        )

        Add-Line "### Authentication policies"
        Add-Line
        $ap = @()
        try { $ap = @(Get-AuthenticationPolicy) } catch { Add-LogEntry -Section 'AuthenticationPolicy' -ErrorRecord $_ }
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
        try { $rg = @(Get-RoleGroup) } catch { Add-LogEntry -Section 'RoleGroup' -ErrorRecord $_ }
        $rgRows = @()
        foreach ($g in $rg) {
            $members = @()
            try { $members = @(Get-RoleGroupMember -Identity $g.Name) } catch { Add-LogEntry -Section "RoleGroupMember ($($g.Name))" -ErrorRecord $_ }
            $info = Get-RoleGroupMemberInfo -Members $members
            $rgRows += [PSCustomObject]@{
                RoleGroup = (Get-RoleGroupDisplayName $g.Name)
                Members = $info.Summary
                GroupMembers = $info.Groups -join '; '
                UserMembers = $info.Users -join '; '
                ManagedBy = @($g.ManagedBy) -join '; '
            }
        }
        Out-Table -CsvName "RoleGroups" -Rows $rgRows -Columns @('RoleGroup','Members','GroupMembers','UserMembers')

        Add-Line "### Role assignment policies (user roles)"
        Add-Line
        $rap = @()
        try { $rap = @(Get-RoleAssignmentPolicy) } catch { Add-LogEntry -Section 'RoleAssignmentPolicy' -ErrorRecord $_ }
        $mbx = @()
        try { $mbx = Get-SharedMailboxes } catch { }
        $rapRows = @($rap | ForEach-Object {
            $rapName = $_.Name
            $count = @($mbx | Where-Object { $_.RoleAssignmentPolicy -eq $rapName }).Count
            [PSCustomObject]@{
                Name = $rapName
                Description = $_.Description
                IsDefault = $_.IsDefault
                AssignedRoles = @($_.AssignedRoles | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.Name) { "$($_.Name)" } else { "$_" } } | Sort-Object) -join '; '
                MailboxCount = $count
            }
        })
        Out-Table -CsvName "RoleAssignmentPolicies" -Rows $rapRows -Columns @('Name','Description','IsDefault','AssignedRoles','MailboxCount')

        Add-Line "### Custom management roles"
        Add-Line
        $customRoles = @()
        try { $customRoles = @(Get-ManagementRole | Where-Object { -not (Test-IsTrue $_.IsEndUserRole) -and -not (Test-IsTrue $_.IsRootRole) } ) } catch { Add-LogEntry -Section 'Custom management roles' -ErrorRecord $_ }
        $crRows = @($customRoles | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Parent = $_.Parent; RoleType = $_.RoleType }
        })
        Out-Table -CsvName "CustomManagementRoles" -Rows $crRows -Columns @('Name','Parent','RoleType')

        Add-Line "### Direct user role assignments"
        Add-Line
        $assignments = @()
        try { $assignments = @(Get-ManagementRoleAssignment -RoleAssigneeType User) } catch { Add-LogEntry -Section 'Direct role assignments' -ErrorRecord $_ }
        $asRows = @($assignments | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Role = $_.Role
                RoleAssignee = $_.RoleAssigneeName
                AssignmentMethod = $_.AssignmentMethod
                Enabled = $_.Enabled
                RecipientReadScope = $_.RecipientReadScope
                RecipientWriteScope = $_.RecipientWriteScope
                ConfigReadScope = $_.ConfigReadScope
                ConfigWriteScope = $_.ConfigWriteScope
                CustomRecipientWriteScope = $_.CustomRecipientWriteScope
                CustomConfigWriteScope = $_.CustomConfigWriteScope
                RecipientAdministrativeUnitScope = $_.RecipientAdministrativeUnitScope
            }
        })
        Out-Table -CsvName "DirectRoleAssignments" -Rows $asRows -Columns @('Name','Role','RoleAssignee','AssignmentMethod','Enabled','RecipientReadScope','RecipientWriteScope','ConfigReadScope','ConfigWriteScope','CustomRecipientWriteScope','CustomConfigWriteScope','RecipientAdministrativeUnitScope')

        Add-Line "### Management scopes"
        Add-Line
        $scopes = @()
        try { $scopes = @(Get-ManagementScope) } catch { Add-LogEntry -Section 'ManagementScope' -ErrorRecord $_ }
        $scRows = @($scopes | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ScopeRestrictionType = $_.ScopeRestrictionType; Exclusive = $_.Exclusive
                              RecipientRoot = $_.RecipientRoot; RecipientFilter = $_.RecipientFilter }
        })
        Out-Table -CsvName "ManagementScopes" -Rows $scRows -Columns @('Name','ScopeRestrictionType','Exclusive','RecipientRoot','RecipientFilter')

        Add-Line "### RBAC for Applications"
        Add-Line
        $sps = @()
        try { $sps = @(Get-ServicePrincipal) } catch { Add-LogEntry -Section 'ServicePrincipal' -ErrorRecord $_ }
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
            catch { Add-LogEntry -Section 'SP role assignments' -ErrorRecord $_ }
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

        $presetRulesFailed = $false
        $eopRules = @()
        try { $eopRules = @(Get-EOPProtectionPolicyRule) } catch { $presetRulesFailed = $true; Add-LogEntry -Section 'EOPProtectionPolicyRule' -ErrorRecord $_ }
        $atpPresetRules = @()
        try { $atpPresetRules = @(Get-ATPProtectionPolicyRule) } catch { $presetRulesFailed = $true; Add-LogEntry -Section 'ATPProtectionPolicyRule' -ErrorRecord $_ }
        $presetRules = @($eopRules) + @($atpPresetRules)
        $builtInRules = @()
        try { $builtInRules = @(Get-ATPBuiltInProtectionRule) } catch { $presetRulesFailed = $true; Add-LogEntry -Section 'ATPBuiltInProtectionRule' -ErrorRecord $_ }

        Add-Line "### Anti-phishing policies"
        Add-Line
        $phish = @()
        try { $phish = @(Get-AntiPhishPolicy) } catch { Add-LogEntry -Section 'AntiPhishPolicy' -ErrorRecord $_ }
        $phishRules = @(); $phishRulesFailed = $false
        try { $phishRules = @(Get-AntiPhishRule) } catch { $phishRulesFailed = $true; Add-LogEntry -Section 'AntiPhishRule' -ErrorRecord $_ }
        $pRows = @($phish | ForEach-Object {
            $pol = $_
            $r = @($phishRules | Where-Object { $_.AntiPhishPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -RuleLookupFailed:($phishRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedUsers = $i.IncludedUsers
                IncludedGroups = $i.IncludedGroups
                IncludedDomains = $i.IncludedDomains
                ExcludedUsers = $i.ExcludedUsers
                ExcludedGroups = $i.ExcludedGroups
                ExcludedDomains = $i.ExcludedDomains
                PhishThresholdLevel = $pol.PhishThresholdLevel
                EnableTargetedUserProtection = $pol.EnableTargetedUserProtection
                TargetedUsersToProtect = $pol.TargetedUsersToProtect
                EnableOrganizationDomainsProtection = $pol.EnableOrganizationDomainsProtection
                EnableTargetedDomainsProtection = $pol.EnableTargetedDomainsProtection
                TargetedDomainsToProtect = $pol.TargetedDomainsToProtect
                PolicyExcludedSenders = $pol.ExcludedSenders
                PolicyExcludedDomains = $pol.ExcludedDomains
                EnableMailboxIntelligence = $pol.EnableMailboxIntelligence
                EnableMailboxIntelligenceProtection = $pol.EnableMailboxIntelligenceProtection
                EnableSpoofIntelligence = $pol.EnableSpoofIntelligence
                TargetedUserProtectionAction = $pol.TargetedUserProtectionAction
                TargetedUserQuarantineTag = $pol.TargetedUserQuarantineTag
                TargetedDomainProtectionAction = $pol.TargetedDomainProtectionAction
                TargetedDomainQuarantineTag = $pol.TargetedDomainQuarantineTag
                MailboxIntelligenceProtectionAction = $pol.MailboxIntelligenceProtectionAction
                MailboxIntelligenceQuarantineTag = $pol.MailboxIntelligenceQuarantineTag
                HonorDmarcPolicy = $pol.HonorDmarcPolicy
                DmarcQuarantineAction = $pol.DmarcQuarantineAction
                DmarcRejectAction = $pol.DmarcRejectAction
                AuthenticationFailAction = $pol.AuthenticationFailAction
                SpoofQuarantineTag = $pol.SpoofQuarantineTag
                EnableFirstContactSafetyTips = $pol.EnableFirstContactSafetyTips
                EnableSimilarUsersSafetyTips = $pol.EnableSimilarUsersSafetyTips
                EnableSimilarDomainsSafetyTips = $pol.EnableSimilarDomainsSafetyTips
                EnableUnusualCharactersSafetyTips = $pol.EnableUnusualCharactersSafetyTips
                EnableUnauthenticatedSender = $pol.EnableUnauthenticatedSender
                EnableViaTag = $pol.EnableViaTag
            }
        })
        Out-Table -CsvName "AntiPhishPolicies" -Rows $pRows -Columns @('Name','Status','Priority','AppliesTo','IncludedUsers','IncludedGroups','IncludedDomains','ExcludedUsers','ExcludedGroups','ExcludedDomains','PhishThresholdLevel','EnableTargetedUserProtection','TargetedUsersToProtect','EnableOrganizationDomainsProtection','EnableTargetedDomainsProtection','TargetedDomainsToProtect','PolicyExcludedSenders','PolicyExcludedDomains','EnableMailboxIntelligence','EnableMailboxIntelligenceProtection','EnableSpoofIntelligence','TargetedUserProtectionAction','TargetedUserQuarantineTag','TargetedDomainProtectionAction','TargetedDomainQuarantineTag','MailboxIntelligenceProtectionAction','MailboxIntelligenceQuarantineTag','HonorDmarcPolicy','DmarcQuarantineAction','DmarcRejectAction','AuthenticationFailAction','SpoofQuarantineTag','EnableFirstContactSafetyTips','EnableSimilarUsersSafetyTips','EnableSimilarDomainsSafetyTips','EnableUnusualCharactersSafetyTips','EnableUnauthenticatedSender','EnableViaTag')

        Add-Line "### Anti-spam inbound policies"
        Add-Line
        $spamIn = @()
        try { $spamIn = @(Get-HostedContentFilterPolicy) } catch { Add-LogEntry -Section 'HostedContentFilterPolicy' -ErrorRecord $_ }
        $spamInRules = @(); $spamInRulesFailed = $false
        try { $spamInRules = @(Get-HostedContentFilterRule) } catch { $spamInRulesFailed = $true; Add-LogEntry -Section 'HostedContentFilterRule' -ErrorRecord $_ }
        $siRows = @($spamIn | ForEach-Object {
            $pol = $_
            $r = @($spamInRules | Where-Object { $_.HostedContentFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -RuleLookupFailed:($spamInRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedUsers = $i.IncludedUsers
                IncludedGroups = $i.IncludedGroups
                IncludedDomains = $i.IncludedDomains
                ExcludedUsers = $i.ExcludedUsers
                ExcludedGroups = $i.ExcludedGroups
                ExcludedDomains = $i.ExcludedDomains
                BulkThreshold = $pol.BulkThreshold
                EnableLanguageBlockList = $pol.EnableLanguageBlockList
                LanguageBlockList = $pol.LanguageBlockList
                EnableRegionBlockList = $pol.EnableRegionBlockList
                RegionBlockList = $pol.RegionBlockList
                SpamAction = $pol.SpamAction
                SpamQuarantineTag = $pol.SpamQuarantineTag
                HighConfidenceSpamAction = $pol.HighConfidenceSpamAction
                HighConfidenceSpamQuarantineTag = $pol.HighConfidenceSpamQuarantineTag
                PhishSpamAction = $pol.PhishSpamAction
                PhishQuarantineTag = $pol.PhishQuarantineTag
                HighConfidencePhishAction = $pol.HighConfidencePhishAction
                HighConfidencePhishQuarantineTag = $pol.HighConfidencePhishQuarantineTag
                BulkSpamAction = $pol.BulkSpamAction
                BulkQuarantineTag = $pol.BulkQuarantineTag
                IntraOrgFilterState = $pol.IntraOrgFilterState
                QuarantineRetentionPeriod = $pol.QuarantineRetentionPeriod
                InlineSafetyTipsEnabled = $pol.InlineSafetyTipsEnabled
                PhishZapEnabled = $pol.PhishZapEnabled
                SpamZapEnabled = $pol.SpamZapEnabled
                AllowedSenders = $pol.AllowedSenders
                AllowedSenderDomains = $pol.AllowedSenderDomains
                BlockedSenders = $pol.BlockedSenders
                BlockedSenderDomains = $pol.BlockedSenderDomains
                IncreaseScoreWithImageLinks = $pol.IncreaseScoreWithImageLinks
                IncreaseScoreWithNumericIps = $pol.IncreaseScoreWithNumericIps
                IncreaseScoreWithRedirectToOtherPort = $pol.IncreaseScoreWithRedirectToOtherPort
                IncreaseScoreWithBizOrInfoUrls = $pol.IncreaseScoreWithBizOrInfoUrls
                MarkAsSpamEmptyMessages = $pol.MarkAsSpamEmptyMessages
                MarkAsSpamJavaScriptInHtml = $pol.MarkAsSpamJavaScriptInHtml
                MarkAsSpamFramesInHtml = $pol.MarkAsSpamFramesInHtml
                MarkAsSpamObjectTagsInHtml = $pol.MarkAsSpamObjectTagsInHtml
                MarkAsSpamEmbedTagsInHtml = $pol.MarkAsSpamEmbedTagsInHtml
                MarkAsSpamFormTagsInHtml = $pol.MarkAsSpamFormTagsInHtml
                MarkAsSpamWebBugsInHtml = $pol.MarkAsSpamWebBugsInHtml
                MarkAsSpamSensitiveWordList = $pol.MarkAsSpamSensitiveWordList
                MarkAsSpamSpfRecordHardFail = $pol.MarkAsSpamSpfRecordHardFail
                MarkAsSpamFromAddressAuthFail = $pol.MarkAsSpamFromAddressAuthFail
                MarkAsSpamNdrBackscatter = $pol.MarkAsSpamNdrBackscatter
                MarkAsSpamBulkMail = $pol.MarkAsSpamBulkMail
                TestModeAction = $pol.TestModeAction
            }
        })
        Out-Table -CsvName "AntiSpamInboundPolicies" -Rows $siRows -Columns @('Name','Status','Priority','AppliesTo','IncludedUsers','IncludedGroups','IncludedDomains','ExcludedUsers','ExcludedGroups','ExcludedDomains','BulkThreshold','EnableLanguageBlockList','LanguageBlockList','EnableRegionBlockList','RegionBlockList','SpamAction','SpamQuarantineTag','HighConfidenceSpamAction','HighConfidenceSpamQuarantineTag','PhishSpamAction','PhishQuarantineTag','HighConfidencePhishAction','HighConfidencePhishQuarantineTag','BulkSpamAction','BulkQuarantineTag','IntraOrgFilterState','QuarantineRetentionPeriod','InlineSafetyTipsEnabled','PhishZapEnabled','SpamZapEnabled','AllowedSenders','AllowedSenderDomains','BlockedSenders','BlockedSenderDomains','IncreaseScoreWithImageLinks','IncreaseScoreWithNumericIps','IncreaseScoreWithRedirectToOtherPort','IncreaseScoreWithBizOrInfoUrls','MarkAsSpamEmptyMessages','MarkAsSpamJavaScriptInHtml','MarkAsSpamFramesInHtml','MarkAsSpamObjectTagsInHtml','MarkAsSpamEmbedTagsInHtml','MarkAsSpamFormTagsInHtml','MarkAsSpamWebBugsInHtml','MarkAsSpamSensitiveWordList','MarkAsSpamSpfRecordHardFail','MarkAsSpamFromAddressAuthFail','MarkAsSpamNdrBackscatter','MarkAsSpamBulkMail','TestModeAction')

        Add-Line "### Connection filter policy"
        Add-Line
        $cf = @()
        try { $cf = @(Get-HostedConnectionFilterPolicy) } catch { Add-LogEntry -Section 'HostedConnectionFilterPolicy' -ErrorRecord $_ }
        $cfRows = @($cf | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Status = 'On (tenant-wide policy)'
                IPAllowList = $_.IPAllowList
                IPBlockList = $_.IPBlockList
                EnableSafeList = $_.EnableSafeList
            }
        })
        Out-Table -CsvName "ConnectionFilterPolicy" -Rows $cfRows -Columns @('Name','Status','IPAllowList','IPBlockList','EnableSafeList')

        Add-Line "### Outbound spam policies"
        Add-Line
        $spamOut = @()
        try { $spamOut = @(Get-HostedOutboundSpamFilterPolicy) } catch { Add-LogEntry -Section 'HostedOutboundSpamFilterPolicy' -ErrorRecord $_ }
        $spamOutRules = @(); $spamOutRulesFailed = $false
        try { $spamOutRules = @(Get-HostedOutboundSpamFilterRule) } catch { $spamOutRulesFailed = $true; Add-LogEntry -Section 'HostedOutboundSpamFilterRule' -ErrorRecord $_ }
        $soRows = @($spamOut | ForEach-Object {
            $pol = $_
            $r = @($spamOutRules | Where-Object { $_.HostedOutboundSpamFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -SenderBased -RuleLookupFailed:($spamOutRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedSenders = $i.IncludedUsers
                IncludedSenderGroups = $i.IncludedGroups
                IncludedSenderDomains = $i.IncludedDomains
                ExcludedSenders = $i.ExcludedUsers
                ExcludedSenderGroups = $i.ExcludedGroups
                ExcludedSenderDomains = $i.ExcludedDomains
                RecipientLimitExternalPerHour = $pol.RecipientLimitExternalPerHour
                RecipientLimitInternalPerHour = $pol.RecipientLimitInternalPerHour
                RecipientLimitPerDay = $pol.RecipientLimitPerDay
                ActionWhenThresholdReached = $pol.ActionWhenThresholdReached
                AutoForwardingMode = $pol.AutoForwardingMode
                BccSuspiciousOutboundMail = $pol.BccSuspiciousOutboundMail
                BccSuspiciousOutboundAdditionalRecipients = $pol.BccSuspiciousOutboundAdditionalRecipients
                NotifyOutboundSpam = $pol.NotifyOutboundSpam
                NotifyOutboundSpamRecipients = $pol.NotifyOutboundSpamRecipients
            }
        })
        Out-Table -CsvName "OutboundSpamPolicies" -Rows $soRows -Columns @('Name','Status','Priority','AppliesTo','IncludedSenders','IncludedSenderGroups','IncludedSenderDomains','ExcludedSenders','ExcludedSenderGroups','ExcludedSenderDomains','RecipientLimitExternalPerHour','RecipientLimitInternalPerHour','RecipientLimitPerDay','ActionWhenThresholdReached','AutoForwardingMode','BccSuspiciousOutboundMail','BccSuspiciousOutboundAdditionalRecipients','NotifyOutboundSpam','NotifyOutboundSpamRecipients')

        Add-Line "### Anti-malware policies"
        Add-Line
        $mal = @()
        try { $mal = @(Get-MalwareFilterPolicy) } catch { Add-LogEntry -Section 'MalwareFilterPolicy' -ErrorRecord $_ }
        $malRules = @(); $malRulesFailed = $false
        try { $malRules = @(Get-MalwareFilterRule) } catch { $malRulesFailed = $true; Add-LogEntry -Section 'MalwareFilterRule' -ErrorRecord $_ }
        $malRows = @($mal | ForEach-Object {
            $pol = $_
            $r = @($malRules | Where-Object { $_.MalwareFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -RuleLookupFailed:($malRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedUsers = $i.IncludedUsers
                IncludedGroups = $i.IncludedGroups
                IncludedDomains = $i.IncludedDomains
                ExcludedUsers = $i.ExcludedUsers
                ExcludedGroups = $i.ExcludedGroups
                ExcludedDomains = $i.ExcludedDomains
                EnableFileFilter = $pol.EnableFileFilter
                FileTypes = $pol.FileTypes
                FileTypeAction = $pol.FileTypeAction
                ZapEnabled = $pol.ZapEnabled
                QuarantineTag = $pol.QuarantineTag
                EnableInternalSenderAdminNotifications = $pol.EnableInternalSenderAdminNotifications
                InternalSenderAdminAddress = $pol.InternalSenderAdminAddress
                EnableExternalSenderAdminNotifications = $pol.EnableExternalSenderAdminNotifications
                ExternalSenderAdminAddress = $pol.ExternalSenderAdminAddress
                CustomNotifications = $pol.CustomNotifications
                CustomFromName = $pol.CustomFromName
                CustomFromAddress = $pol.CustomFromAddress
            }
        })
        Out-Table -CsvName "MalwareFilterPolicies" -Rows $malRows -Columns @('Name','Status','Priority','AppliesTo','IncludedUsers','IncludedGroups','IncludedDomains','ExcludedUsers','ExcludedGroups','ExcludedDomains','EnableFileFilter','FileTypes','FileTypeAction','ZapEnabled','QuarantineTag','EnableInternalSenderAdminNotifications','InternalSenderAdminAddress','EnableExternalSenderAdminNotifications','ExternalSenderAdminAddress','CustomNotifications','CustomFromName','CustomFromAddress')

        Add-Line "### Safe Attachments policies"
        Add-Line
        $sa = @()
        try { $sa = @(Get-SafeAttachmentPolicy) } catch { Add-LogEntry -Section 'SafeAttachmentPolicy' -ErrorRecord $_ }
        $saRules = @(); $saRulesFailed = $false
        try { $saRules = @(Get-SafeAttachmentRule) } catch { $saRulesFailed = $true; Add-LogEntry -Section 'SafeAttachmentRule' -ErrorRecord $_ }
        $saRows = @($sa | ForEach-Object {
            $pol = $_
            $r = @($saRules | Where-Object { $_.SafeAttachmentPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -RuleLookupFailed:($saRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedUsers = $i.IncludedUsers
                IncludedGroups = $i.IncludedGroups
                IncludedDomains = $i.IncludedDomains
                ExcludedUsers = $i.ExcludedUsers
                ExcludedGroups = $i.ExcludedGroups
                ExcludedDomains = $i.ExcludedDomains
                Enable = $pol.Enable
                Action = $pol.Action
                QuarantineTag = $pol.QuarantineTag
                Redirect = $pol.Redirect
                RedirectAddress = $pol.RedirectAddress
                EnableBlockingEncryptedAttachments = $pol.EnableBlockingEncryptedAttachments
                ExcludedTypesFromBlockingEncryptedAttachments = $pol.ExcludedTypesFromBlockingEncryptedAttachments
                QuarantineTagForBlockingEncryptedAttachments = $pol.QuarantineTagForBlockingEncryptedAttachments
            }
        })
        Out-Table -CsvName "SafeAttachmentPolicies" -Rows $saRows -Columns @('Name','Status','Priority','AppliesTo','IncludedUsers','IncludedGroups','IncludedDomains','ExcludedUsers','ExcludedGroups','ExcludedDomains','Enable','Action','QuarantineTag','Redirect','RedirectAddress','EnableBlockingEncryptedAttachments','ExcludedTypesFromBlockingEncryptedAttachments','QuarantineTagForBlockingEncryptedAttachments')

        Add-Line "### ATP policy for O365"
        Add-Line
        $atp = @()
        try { $atp = @(Get-AtpPolicyForO365) } catch { Add-LogEntry -Section 'AtpPolicyForO365' -ErrorRecord $_ }
        $atpRows = @($atp | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                EnableATPForSPOTeamsODB = $_.EnableATPForSPOTeamsODB
                EnableSafeDocs = $_.EnableSafeDocs
                AllowSafeDocsOpen = $_.AllowSafeDocsOpen
            }
        })
        Out-Table -CsvName "AtpPolicyForO365" -Rows $atpRows -Columns @('Name','EnableATPForSPOTeamsODB','EnableSafeDocs','AllowSafeDocsOpen')

        Add-Line "### Safe Links policies"
        Add-Line
        $sl = @()
        try { $sl = @(Get-SafeLinksPolicy) } catch { Add-LogEntry -Section 'SafeLinksPolicy' -ErrorRecord $_ }
        $slRules = @(); $slRulesFailed = $false
        try { $slRules = @(Get-SafeLinksRule) } catch { $slRulesFailed = $true; Add-LogEntry -Section 'SafeLinksRule' -ErrorRecord $_ }
        $slRows = @($sl | ForEach-Object {
            $pol = $_
            $r = @($slRules | Where-Object { $_.SafeLinksPolicy -eq $pol.Name }) | Select-Object -First 1
            $i = Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtInRules -RuleLookupFailed:($slRulesFailed -or $presetRulesFailed)
            [PSCustomObject]@{
                Name = $pol.Name
                Status = $i.Status
                Priority = $i.Priority
                AppliesTo = $i.AppliesTo
                IncludedUsers = $i.IncludedUsers
                IncludedGroups = $i.IncludedGroups
                IncludedDomains = $i.IncludedDomains
                ExcludedUsers = $i.ExcludedUsers
                ExcludedGroups = $i.ExcludedGroups
                ExcludedDomains = $i.ExcludedDomains
                EnableSafeLinksForEmail = $pol.EnableSafeLinksForEmail
                EnableForInternalSenders = $pol.EnableForInternalSenders
                ScanUrls = $pol.ScanUrls
                DeliverMessageAfterScan = $pol.DeliverMessageAfterScan
                DisableUrlRewrite = $pol.DisableUrlRewrite
                DoNotRewriteUrls = $pol.DoNotRewriteUrls
                EnableSafeLinksForTeams = $pol.EnableSafeLinksForTeams
                EnableSafeLinksForOffice = $pol.EnableSafeLinksForOffice
                TrackClicks = $pol.TrackClicks
                AllowClickThrough = $pol.AllowClickThrough
                EnableOrganizationBranding = $pol.EnableOrganizationBranding
                CustomNotificationText = $pol.CustomNotificationText
                UseTranslatedNotificationText = $pol.UseTranslatedNotificationText
            }
        })
        Out-Table -CsvName "SafeLinksPolicies" -Rows $slRows -Columns @('Name','Status','Priority','AppliesTo','IncludedUsers','IncludedGroups','IncludedDomains','ExcludedUsers','ExcludedGroups','ExcludedDomains','EnableSafeLinksForEmail','EnableForInternalSenders','ScanUrls','DeliverMessageAfterScan','DisableUrlRewrite','DoNotRewriteUrls','EnableSafeLinksForTeams','EnableSafeLinksForOffice','TrackClicks','AllowClickThrough','EnableOrganizationBranding','CustomNotificationText','UseTranslatedNotificationText')

        Add-Line "### Preset security policies"
        Add-Line
        $presetRows = @()
        try {
            $presetRows += @($eopRules | ForEach-Object {
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
        } catch { Add-LogEntry -Section 'EOPProtectionPolicyRule' -ErrorRecord $_ }
        try {
            $presetRows += @($atpPresetRules | ForEach-Object {
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
        } catch { Add-LogEntry -Section 'ATPProtectionPolicyRule' -ErrorRecord $_ }
        Out-Table -CsvName "PresetSecurityPolicies" -Rows $presetRows -Columns @('Name','State','Priority','Type','SentTo','SentToMemberOf','RecipientDomainIs')

        Add-Line "### Built-in protection rule"
        Add-Line
        $biRows = @($builtInRules | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; State = $_.State; SentTo = $_.SentTo; SentToMemberOf = $_.SentToMemberOf; RecipientDomainIs = $_.RecipientDomainIs }
        })
        Out-Table -CsvName "BuiltInProtectionRule" -Rows $biRows -Columns @('Name','State','SentTo','SentToMemberOf','RecipientDomainIs')

        Add-Line "### Tenant Allow/Block List"
        Add-Line
        $tablLists = [ordered]@{ 'Sender' = 'Domains & addresses'; 'Url' = 'URLs'; 'FileHash' = 'Files'; 'IP' = 'IP addresses' }
        $tablRows = @()
        foreach ($listType in $tablLists.Keys) {
            try {
                $tablRows += @(Get-TenantAllowBlockListItems -ListType $listType -ErrorAction Stop | ForEach-Object {
                    [PSCustomObject]@{
                        List = $tablLists[$listType]
                        Value = $_.Value
                        Action = $_.Action
                        ExpirationDate = $(if ($_.ExpirationDate) { $_.ExpirationDate } else { 'Never expire' })
                        Notes = $_.Notes
                    }
                })
            }
            catch { Add-LogEntry -Section "TenantAllowBlockListItems ($listType)" -ErrorRecord $_ }
        }
        Out-Table -CsvName "TenantAllowBlockList" -Rows $tablRows -Columns @('List','Value','Action','ExpirationDate','Notes')
        $spoof = @()
        try { $spoof = @(Get-TenantAllowBlockListSpoofItems) } catch { Add-LogEntry -Section 'TenantAllowBlockListSpoofItems' -ErrorRecord $_ }
        $spoofRows = @($spoof | ForEach-Object {
            [PSCustomObject]@{ SpoofedUser = $_.SpoofedUser; SendingInfrastructure = $_.SendingInfrastructure; Action = $_.Action; SpoofType = $_.SpoofType }
        })
        Out-Table -CsvName "TenantAllowBlockListSpoof" -Rows $spoofRows -Columns @('SpoofedUser','SendingInfrastructure','Action','SpoofType')

        Add-Line "### Advanced delivery (SecOps mailboxes and phishing simulations)"
        Add-Line
        $secOps = @()
        try { $secOps = @(Get-SecOpsOverridePolicy) } catch { Add-LogEntry -Section 'SecOpsOverridePolicy' -ErrorRecord $_ }
        $secOpsRules = @(); $secOpsRulesFailed = $false
        try { $secOpsRules = @(Get-ExoSecOpsOverrideRule) } catch { $secOpsRulesFailed = $true; Add-LogEntry -Section 'ExoSecOpsOverrideRule' -ErrorRecord $_ }
        $secOpsRows = @($secOps | ForEach-Object {
            $pol = $_
            $ids = @("$($pol.Identity)", "$($pol.Name)", "$($pol.Guid)", "$($pol.ExchangeObjectId)", "$($pol.Id)", "$($pol.DistinguishedName)") | Where-Object { $_ }
            $r = @($secOpsRules | Where-Object { "$($_.Policy)" -in $ids }) | Select-Object -First 1
            if (-not $r -and $secOps.Count -eq 1 -and $secOpsRules.Count -ge 1) { $r = $secOpsRules[0] }
            [PSCustomObject]@{ Name = $pol.Name; SentTo = $(if ($secOpsRulesFailed) { 'Unknown (rule lookup failed)' } else { $r.SentTo }); SentToMemberOf = $r.SentToMemberOf }
        })
        Out-Table -CsvName "SecOpsOverridePolicy" -Rows $secOpsRows -Columns @('Name','SentTo','SentToMemberOf')
        $phishSim = @()
        try { $phishSim = @(Get-PhishSimOverridePolicy) } catch { Add-LogEntry -Section 'PhishSimOverridePolicy' -ErrorRecord $_ }
        $phishSimRules = @(); $phishSimRulesFailed = $false
        try { $phishSimRules = @(Get-ExoPhishSimOverrideRule) } catch { $phishSimRulesFailed = $true; Add-LogEntry -Section 'ExoPhishSimOverrideRule' -ErrorRecord $_ }
        $simRows = @($phishSim | ForEach-Object {
            $pol = $_
            $ids = @("$($pol.Identity)", "$($pol.Name)", "$($pol.Guid)", "$($pol.ExchangeObjectId)", "$($pol.Id)", "$($pol.DistinguishedName)") | Where-Object { $_ }
            $r = @($phishSimRules | Where-Object { "$($_.Policy)" -in $ids }) | Select-Object -First 1
            if (-not $r -and $phishSim.Count -eq 1 -and $phishSimRules.Count -ge 1) { $r = $phishSimRules[0] }
            [PSCustomObject]@{ Name = $pol.Name; Domains = $(if ($phishSimRulesFailed) { 'Unknown (rule lookup failed)' } else { $r.Domains }); SenderIpRanges = $r.SenderIpRanges }
        })
        Out-Table -CsvName "PhishSimOverridePolicy" -Rows $simRows -Columns @('Name','Domains','SenderIpRanges')
        $simUrls = @()
        try { $simUrls = @(Get-TenantAllowBlockListItems -ListType Url -ListSubType AdvancedDelivery) } catch { Add-LogEntry -Section 'Simulation URLs' -ErrorRecord $_ }
        $simUrlRows = @($simUrls | ForEach-Object {
            [PSCustomObject]@{ SimulationUrl = $_.Value }
        })
        Out-Table -CsvName "SimulationUrls" -Rows $simUrlRows -Columns @('SimulationUrl')

        Add-Line "### Quarantine policies"
        Add-Line
        $qp = @()
        try { $qp = @(Get-QuarantinePolicy) } catch { Add-LogEntry -Section 'QuarantinePolicy' -ErrorRecord $_ }
        $qpRows = @($qp | ForEach-Object {
            $qi = Get-QuarantinePermissionInfo -Policy $_
            [PSCustomObject]@{
                Name = $_.Name
                RecipientMessageAccess = $qi.Access
                Permissions = $qi.Permissions
                QuarantineNotification = $qi.Notifications
                IncludeMessagesFromBlockedSenderAddress = $_.IncludeMessagesFromBlockedSenderAddress
                EndUserQuarantinePermissionsValue = $_.EndUserQuarantinePermissionsValue
            }
        })
        Out-Table -CsvName "QuarantinePolicies" -Rows $qpRows -Columns @('Name','RecipientMessageAccess','Permissions','QuarantineNotification','IncludeMessagesFromBlockedSenderAddress','EndUserQuarantinePermissionsValue')
        $qgs = @()
        try { $qgs = @(Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy) } catch { }
        $qgRows = @($qgs | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; EndUserSpamNotificationFrequency = $_.EndUserSpamNotificationFrequency; ESNEnabled = $_.ESNEnabled }
        })
        Out-Table -CsvName "GlobalQuarantineSettings" -Rows $qgRows -Columns @('Name','EndUserSpamNotificationFrequency','ESNEnabled')

        Add-Line "### Priority account protection and reporting"
        Add-Line
        $ets = @()
        try { $ets = @(Get-EmailTenantSettings) } catch { Add-LogEntry -Section 'EmailTenantSettings' -ErrorRecord $_ }
        $etsRows = @($ets | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; EnablePriorityAccountProtection = $_.EnablePriorityAccountProtection }
        })
        Out-Table -CsvName "EmailTenantSettings" -Rows $etsRows -Columns @('Identity','EnablePriorityAccountProtection')
        $rsp = @()
        try { $rsp = @(Get-ReportSubmissionPolicy) } catch { Add-LogEntry -Section 'ReportSubmissionPolicy' -ErrorRecord $_ }
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
        try { $tpp = @(Get-TeamsProtectionPolicy) } catch { Add-LogEntry -Section 'TeamsProtectionPolicy' -ErrorRecord $_ }
        $tppRows = @($tpp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ZapEnabled = $_.ZapEnabled; HighConfidencePhishQuarantineTag = $_.HighConfidencePhishQuarantineTag; MalwareQuarantineTag = $_.MalwareQuarantineTag }
        })
        Out-Table -CsvName "TeamsProtectionPolicy" -Rows $tppRows -Columns @('Name','ZapEnabled','HighConfidencePhishQuarantineTag','MalwareQuarantineTag')
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
                Type = $(if (Test-IsTrue $_.IsSystemRule) { 'System' } else { 'Custom' })
                Status = $(if (Test-IsTrue $_.Disabled) { 'Off' } else { 'On' })
                Category = $_.Category
                Severity = $_.Severity
                NotifyUser = $_.NotifyUser
                ThreatType = $_.ThreatType
                ManagedIn = $(if ($_.Category -eq 'DataLossPrevention') { 'Purview DLP policy' } else { 'Defender portal (Alert policy)' })
            }
        })
        Out-Table -CsvName "ProtectionAlerts" -Rows $rows -Columns @('Name','Type','Status','Category','Severity','NotifyUser','ThreatType','ManagedIn')
        $script:Summary['Alert policies'] = $alerts.Count
    }
}

function Get-ComplianceSection {
    Invoke-Section -Title $script:SectionTitles.Compliance -Body {
        Add-Line "### MRM retention policies and tags"
        Add-Line
        $rp = @()
        try { $rp = @(Get-RetentionPolicy) } catch { Add-LogEntry -Section 'RetentionPolicy' -ErrorRecord $_ }
        $rpRows = @($rp | ForEach-Object {
            [PSCustomObject]@{
                Name = $(if ($_.Name -eq 'ArbitrationMailbox') { 'ArbitrationMailbox (system policy, hidden in the portal)' } else { $_.Name })
                RetentionPolicyTagLinks = @($_.RetentionPolicyTagLinks | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.Name) { "$($_.Name)" } else { "$_" } } | Sort-Object) -join '; '
                IsDefault = $_.IsDefault
            }
        })
        Out-Table -CsvName "RetentionPolicies" -Rows $rpRows -Columns @('Name','RetentionPolicyTagLinks','IsDefault')
        $tags = @()
        try { $tags = @(Get-RetentionPolicyTag) } catch { Add-LogEntry -Section 'RetentionPolicyTag' -ErrorRecord $_ }
        $tagRows = @($tags | Where-Object { -not $_.SystemTag } | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Type = $_.Type
                RetentionPeriod = Format-RetentionPeriod -Tag $_
                RetentionAction = $_.RetentionAction
                RetentionEnabled = $_.RetentionEnabled
                AgeLimitForRetention = $_.AgeLimitForRetention
            }
        } | Sort-Object Type, Name)
        Out-Table -CsvName "RetentionPolicyTags" -Rows $tagRows -Columns @('Name','Type','RetentionPeriod','RetentionAction','RetentionEnabled','AgeLimitForRetention')

        Add-Line "### IRM and OME configuration"
        Add-Line
        $irm = @()
        try { $irm = @(Get-IRMConfiguration) } catch { Add-LogEntry -Section 'IRMConfiguration' -ErrorRecord $_ }
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
        try { $ome = @(Get-OMEConfiguration) } catch { Add-LogEntry -Section 'OMEConfiguration' -ErrorRecord $_ }
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
        $tagNames = @{}
        try {
            foreach ($t in @(Get-ComplianceTag)) {
                foreach ($k in @("$($t.Guid)", "$($t.ImmutableId)", "$($t.Identity)", "$($t.Name)")) {
                    if ($k -and -not $tagNames.ContainsKey($k)) { $tagNames[$k] = "$($t.Name)" }
                }
            }
        }
        catch { Add-LogEntry -Section 'ComplianceTag (name map)' -ErrorRecord $_ }
        $exo = @($policies | Where-Object {
            @($_.ExchangeLocation).Count -gt 0 -and "$($_.ExchangeLocation)" -notmatch '^\s*$'
        })
        $rows = @($exo | ForEach-Object {
            $pn = $_.Name
            $rules = @()
            try { $rules = @(@(Get-RetentionComplianceRule -Policy $pn) | ForEach-Object { Format-RetentionRule -Rule $_ -TagNames $tagNames } | Sort-Object -Unique) }
            catch { Add-LogEntry -Section "RetentionComplianceRule ($pn)" -ErrorRecord $_; $rules = @('Unknown (rule lookup failed)') }
            [PSCustomObject]@{
                Name = $pn
                Mode = $_.Mode
                Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation
                ExchangeLocationException = $_.ExchangeLocationException
                RetentionSettings = $rules -join '; '
            }
        })
        Out-Table -CsvName "PurviewRetentionPolicies" -Rows $rows -Columns @('Name','Mode','Enabled','ExchangeLocation','ExchangeLocationException','RetentionSettings')
        $script:Summary['Purview retention policies (Exchange)'] = $exo.Count
    }

    Invoke-Section -Title "DLP policies covering Exchange" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'DLP policies') { return }
        Assert-Cmdlet Get-DlpCompliancePolicy
        $policies = @(Get-DlpCompliancePolicy)
        $exo = @($policies | Where-Object { @($_.ExchangeLocation).Count -gt 0 })
        $rows = @($exo | ForEach-Object {
            $pn = $_.Name
            $rules = @(); $rulesFailed = $false
            try { $rules = @(Get-DlpComplianceRule -Policy $pn) } catch { $rulesFailed = $true; Add-LogEntry -Section "DlpComplianceRule ($pn)" -ErrorRecord $_ }
            [PSCustomObject]@{
                Name = $pn
                Mode = $_.Mode
                Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation
                Rules = $(if ($rulesFailed) { 'Unknown (rule lookup failed)' } else { @($rules | ForEach-Object { $_.Name }) -join '; ' })
            }
        })
        Out-Table -CsvName "DlpPolicies" -Rows $rows -Columns @('Name','Mode','Enabled','ExchangeLocation','Rules')
        $script:Summary['DLP policies (Exchange)'] = $exo.Count
    }

    Invoke-Section -Title "Sensitivity labels and label policies" -Level 3 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Sensitivity labels') { return }
        Assert-Cmdlet Get-Label
        $labels = @(Get-Label)
        $labelRows = @($labels | Sort-Object Priority | ForEach-Object {
            $li = Get-SensitivityLabelInfo -Label $_ -AllLabels $labels
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                Name = $_.Name
                Priority = $_.Priority
                ParentLabel = $li.Parent
                Scope = $li.Scope
                DescriptionForUsers = $_.Tooltip
                AccessControl = $li.AccessControl
                ContentMarking = $li.ContentMarking
                AutoLabeling = $li.AutoLabeling
                GroupSettings = $li.GroupSettings
                SiteSettings = $li.SiteSettings
            }
        })
        Out-Table -CsvName "SensitivityLabels" -Rows $labelRows -Columns @('DisplayName','Name','Priority','ParentLabel','Scope','DescriptionForUsers','AccessControl','ContentMarking','AutoLabeling','GroupSettings','SiteSettings')
        $script:Summary['Sensitivity labels'] = $labels.Count

        $labelPolicies = @()
        try { $labelPolicies = @(Get-LabelPolicy) } catch { Add-LogEntry -Section 'LabelPolicy' -ErrorRecord $_ }
        $lpRows = @($labelPolicies | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Description = $_.Comment
                PublishedLabels = @(Resolve-LabelNames -Ids @($_.Labels) -AllLabels $labels) -join '; '
                PublishedTo = @(@($_.ExchangeLocation) + @($_.ModernGroupLocation) | Where-Object { "$_" -match '\S' }) -join '; '
                ExcludedUsersAndGroups = @(@($_.ExchangeLocationException) + @($_.ModernGroupLocationException) | Where-Object { "$_" -match '\S' }) -join '; '
                ModernGroupLocation = @($_.ModernGroupLocation) -join '; '
                ModernGroupLocationException = @($_.ModernGroupLocationException) -join '; '
            }
        })
        Out-Table -CsvName "LabelPolicies" -Rows $lpRows -Columns @('Name','Description','PublishedLabels','PublishedTo','ExcludedUsersAndGroups')
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

if ($ExoLogPath -and -not (Test-Path -LiteralPath $ExoLogPath)) {
    New-Item -ItemType Directory -Force -Path $ExoLogPath | Out-Null
}

try {
    Initialize-ExchangeOnlineModule

    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
    $script:PreExistingConnectionIds = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ConnectionId })
    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($UserPrincipalName) { $exoParams.UserPrincipalName = $UserPrincipalName }
    if ($AppId) {
        $exoParams.AppId = $AppId
        $exoParams.CertificateThumbprint = $CertificateThumbprint
        $exoParams.Organization = $Organization
    }
    if ($ExoLogPath) {
        $exoParams.EnableErrorReporting = $true
        $exoParams.LogDirectoryPath = $ExoLogPath
        $exoParams.LogLevel = 'All'
    }
    try {
        Connect-ExchangeOnline @exoParams
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        $messages = @(); $ex = $_.Exception
        while ($ex) { $messages += $ex.Message; $ex = $ex.InnerException }
        Write-Error "Failed to connect to Exchange Online: $(($messages | Select-Object -Unique) -join ' --> ')"
        exit 1
    }

    # Pre-collect tenant identity for the header
    try {
        $script:OrgConfig = Get-OrganizationConfig -ErrorAction Stop
        $script:TenantName = $script:OrgConfig.DisplayName
    }
    catch { Add-LogEntry -Section 'OrganizationConfig (header)' -ErrorRecord $_ }
    try {
        $script:AcceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop)
        $script:InitialDomain = @($script:AcceptedDomains | Where-Object { $_.InitialDomain } | Select-Object -First 1).DomainName
    }
    catch { Add-LogEntry -Section 'AcceptedDomain (header)' -ErrorRecord $_ }

    # Collected-by identity
    try {
        $connInfo = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
        if ($connInfo.UserPrincipalName) { $script:CollectedBy = $connInfo.UserPrincipalName }
        elseif ($connInfo.User) { $script:CollectedBy = $connInfo.User }
    }
    catch { }

    if ($IncludePurview) {
        Write-Host "Connecting to Security & Compliance PowerShell (Purview)..." -ForegroundColor Cyan
        $ippsParams = @{
            ShowBanner  = $false
            ErrorAction = 'Stop'
            CommandName = @('Get-ProtectionAlert', 'Get-RetentionCompliancePolicy', 'Get-RetentionComplianceRule',
                'Get-DlpCompliancePolicy', 'Get-DlpComplianceRule', 'Get-Label', 'Get-LabelPolicy', 'Get-ComplianceTag')
        }
        if ($UserPrincipalName) { $ippsParams.UserPrincipalName = $UserPrincipalName }
        if ($AppId) {
            $ippsParams.AppId = $AppId
            $ippsParams.CertificateThumbprint = $CertificateThumbprint
            $ippsParams.Organization = $Organization
        }
        if ($ExoLogPath) {
            $ippsParams.EnableErrorReporting = $true
            $ippsParams.LogDirectoryPath = $ExoLogPath
            $ippsParams.LogLevel = 'All'
        }
        try {
            Connect-IPPSSession @ippsParams
            $script:PurviewConnected = $true
            Write-Host "Connected to Security & Compliance." -ForegroundColor Green
        }
        catch {
            $script:PurviewError = Get-ErrorDetail $_
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
                $script:GraphError = Get-ErrorDetail $_
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

    Register-ExoProxyWrapper

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
            @(ConvertTo-CsvRow -Rows $script:CsvData[$key]) | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "CSV exports written to: $csvDir" -ForegroundColor Green
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Host "`nDisconnecting sessions opened by this script..." -ForegroundColor Cyan
    $current = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.ConnectionId })
    $ours = @($current | Where-Object { $script:PreExistingConnectionIds -notcontains $_.ConnectionId })
    if ($ours.Count -gt 0) {
        foreach ($c in $ours) {
            Disconnect-ExchangeOnline -ConnectionId $c.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    elseif ($current.Count -eq 0 -and $script:PreExistingConnectionIds.Count -eq 0) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    if ($script:GraphConnected) { Disconnect-MgGraph -ErrorAction SilentlyContinue }
    Write-Host "Done." -ForegroundColor Green
}
