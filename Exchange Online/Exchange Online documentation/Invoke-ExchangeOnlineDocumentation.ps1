<#
.SYNOPSIS
    Documents an Exchange Online tenant as readable Markdown plus full CSV exports.

.DESCRIPTION
    Connects to Exchange Online and documents the current Exchange Online,
    Exchange Online Protection (EOP) and Defender for Office 365 configuration
    as an English "as-built" document. Optionally includes Microsoft Purview
    (Security & Compliance PowerShell) data with -IncludePurview.

    The output is a Word-friendly Markdown file (top-level '#' sections, '##'
    subsections, '###' per-object cards, tables of at most 4 columns) plus a
    CSV folder holding every collected dataset in full. The Markdown is meant
    to be converted to Word with pandoc and the company template.

    The script never calls Set-/New-/Remove-/Enable-/Disable- cmdlets.

.PARAMETER CustomerName
    Name of the customer/tenant being documented. Used in the file names.
    Mandatory.

.PARAMETER DocumentDate
    Documentation date shown in the document information section. Defaults
    to today.

.PARAMETER Author
    Name of the person producing the documentation. Optional.

.PARAMETER OutputPath
    Folder for the Markdown document and the CSV folder. Defaults to
    .\Reports (relative to the current directory). Created if missing.

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
    adds the Purview data (alert policies, Purview retention, DLP,
    sensitivity labels, retention labels). Off by default.

.PARAMETER IncludeIntegratedApps
    Collects centrally deployed add-ins (Microsoft 365 admin center >
    Settings > Integrated apps) via Get-OrganizationAddIn. Requires the
    O365CentralizedAddInDeployment module and a separate interactive
    sign-in; not supported with -AppId. Off by default.

.PARAMETER MaxListItems
    Maximum number of items shown when a multi-valued property is rendered
    in-line. The rest is summarised with a "see CSV" note. Default 10.

.PARAMETER MaxTableRows
    Maximum number of rows shown per overview table. Longer lists are
    truncated with a "see CSV" note. Default 25.

.PARAMETER CsvDelimiter
    CSV delimiter override. When omitted, Export-Csv -UseCulture is used so
    Danish Excel opens the files correctly.

.PARAMETER DnsServer
    Optional DNS server for Resolve-DnsName lookups.

.PARAMETER ExoLogPath
    Optional folder for ExchangeOnlineManagement client logs. When set,
    Connect-ExchangeOnline and Connect-IPPSSession run with
    -EnableErrorReporting -LogDirectoryPath <path> -LogLevel All. Useful
    when cmdlets fail with the generic 'server side error' message.

.EXAMPLE
    .\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso"
    EXO-only documentation. Purview data is not collected.

.EXAMPLE
    .\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso" -IncludePurview
    Full documentation including Purview data.

.EXAMPLE
    .\Invoke-ExchangeOnlineDocumentation.ps1 -CustomerName "Contoso" -AppId $id -CertificateThumbprint $thumb -Organization contoso.onmicrosoft.com
    App-only certificate authentication.

.NOTES
    Requires: ExchangeOnlineManagement module (v3)
    Install:    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    Version:    1.2
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CustomerName,

    [Parameter(Mandatory = $false)]
    [datetime]$DocumentDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [string]$Author,

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
    [switch]$IncludeIntegratedApps,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$MaxListItems = 10,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$MaxTableRows = 25,

    [Parameter(Mandatory = $false)]
    [ValidateLength(1, 1)]
    [string]$CsvDelimiter,

    [Parameter(Mandatory = $false)]
    [string]$DnsServer,

    [Parameter(Mandatory = $false)]
    [string]$ExoLogPath
)

$script:ScriptVersion = "1.2"
$script:RunStart = Get-Date
$script:CollectionLog = [System.Collections.Generic.List[object]]::new()
$script:CsvData = [ordered]@{}
$script:CsvIndex = [ordered]@{}
$script:Report = [System.Text.StringBuilder]::new()
$script:PurviewConnected = $false
$script:PurviewError = $null
$script:Mailboxes = $null
$script:CasMailboxes = $null
$script:TenantName = $null
$script:InitialDomain = $null
$script:CollectedBy = $UserPrincipalName
$script:TenantId = $null
$script:OrgConfig = $null
$script:AcceptedDomains = $null
$script:DkimConfig = $null
$script:DistGroups = $null
$script:DynGroups = $null
$script:M365Groups = $null
$script:InboundConnectors = $null
$script:OutboundConnectors = $null
$script:TransportRules = $null
$script:TransportConfig = $null
$script:SpCache = $null
$script:SpCacheFailed = $false
$script:LookupFailedText = 'Unknown (lookup failed)'
$script:IntegratedAppsConnected = $false
$script:IntegratedAppsError = $null
$script:IntegratedAppsList = $null
$script:IntegratedAppsListError = $null
$script:IntegratedAppsDetails = @{}
$script:IntegratedAppsDetailErrors = @{}
$script:MailboxesError = $null
$script:CasMailboxesError = $null
$script:DistGroupsError = $null
$script:DynGroupsError = $null
$script:M365GroupsError = $null
$script:InboundConnectorsError = $null
$script:OutboundConnectorsError = $null
$script:TransportRulesError = $null
$script:TransportConfigError = $null

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
    $fqid = "$($ErrorRecord.FullyQualifiedErrorId)"
    if ($fqid -and -not $text.Contains($fqid)) { $text += " [$fqid]" }
    return $text
}

function Invoke-ExoProxy {
    $pass = @()
    $errAction = $null
    for ($i = 1; $i -lt $args.Count; $i++) {
        if ($args[$i] -is [string] -and $args[$i] -match '^-ErrorAction(:(.*))?$') {
            if ($Matches[2]) { $errAction = $Matches[2] } else { $i++; $errAction = "$($args[$i])" }
            continue
        }
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
    $isBenign403 = {
        param($rec)
        $x = $rec.Exception; $benign = $false
        while ($x) {
            if ($x -is [System.Net.WebException]) { $benign = ($x.Response -and [int]$x.Response.StatusCode -eq 403); break }
            $x = $x.InnerException
        }
        return $benign
    }
    if ($errAction -in 'SilentlyContinue', 'Ignore', 'Continue') { return $out }
    $final = @($records | Where-Object { -not (& $isBenign403 $_) })
    if ($final.Count -gt 0) {
        throw ((@($records) | ForEach-Object { Get-ErrorDetail $_ } | Select-Object -Unique) -join ' | ')
    }
    $benign403 = @($records | Where-Object { & $isBenign403 $_ })
    if ($benign403.Count -gt 0) {
        if (@($out).Count -eq 0) {
            throw ((@($records) | ForEach-Object { Get-ErrorDetail $_ } | Select-Object -Unique) -join ' | ')
        }
        Add-LogEntry -Section $args[0] -Reason ("Partial result (HTTP 403): " + ((@($benign403) | ForEach-Object { Get-ErrorDetail $_ } | Select-Object -Unique) -join ' | '))
    }
    $out
}

function Register-ExoProxyWrapper {
    param([string[]]$ExcludeConnectionIds)
    $script:ExoProxyCommands = @{}
    $names = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $ExcludeConnectionIds -notcontains $_.ConnectionId } | ForEach-Object { $_.ModuleName } | Where-Object { $_ })
    $candidates = @($names) + @($names | ForEach-Object { Split-Path $_ -Leaf }) + @($names | ForEach-Object { [IO.Path]::GetFileNameWithoutExtension($_) })
    $modules = @(Get-Module | Where-Object { $candidates -contains $_.Name -or $candidates -contains $_.Path -or $candidates -contains $_.ModuleBase } | Sort-Object Name -Unique)
    if ($modules.Count -eq 0) {
        if (@($ExcludeConnectionIds).Count -eq 0) {
            $modules = @(Get-Module | Where-Object { $_.Name -like 'tmpEXO_*' } | Sort-Object Name -Unique)
        }
        else {
            throw "Could not identify the Exchange Online session opened by this script. Disconnect other Exchange Online / Security & Compliance sessions (Disconnect-ExchangeOnline) and run again."
        }
    }
    foreach ($mod in $modules) {
        foreach ($cmdName in @($mod.ExportedFunctions.Keys)) {
            $script:ExoProxyCommands[$cmdName] = "$($mod.Name)\$cmdName"
            Set-Item -Path "function:script:$cmdName" -Value ([scriptblock]::Create("Invoke-ExoProxy '$cmdName' @args"))
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
    if ("$Id" -notmatch '\S') { return 'Unnamed member' }
    if ($null -eq $script:DirObjCache) { $script:DirObjCache = @{} }
    if ($script:DirObjCache.ContainsKey($Id)) { return $script:DirObjCache[$Id] }
    if ($null -eq $script:SpCache) { $script:SpCache = @(); try { $script:SpCache = @(Get-ServicePrincipal) } catch { $script:SpCacheFailed = $true } }
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
            if ($v -is [string] -and $v -match '^[=+\-@\t\r]') { $v = "'" + $v }
            $o[$p.Name] = $v
        }
        [PSCustomObject]$o
    }
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
    if ($types -contains 'applycontentmarking' -or $types -contains 'applycontentmarkingheader' -or (Test-IsTrue $Label.ApplyContentMarkingHeaderEnabled)) { $marking += 'Header' }
    if ($types -contains 'applycontentmarkingfooter' -or (Test-IsTrue $Label.ApplyContentMarkingFooterEnabled)) { $marking += 'Footer' }
    if ($types -contains 'applywatermarking' -or (Test-IsTrue $Label.ApplyWaterMarkingEnabled)) { $marking += 'Watermark' }
    $encrypt = ($types -contains 'encrypt') -or (Test-IsTrue $Label.EncryptionEnabled)
    $groupSite = Test-IsTrue $Label.SiteAndGroupProtectionEnabled
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
        if (-not $label) {
            $guidId = "$(@($m.ExternalDirectoryObjectId, $m.Name, $m.Identity) | Where-Object { "$_" -match $guid } | Select-Object -First 1)"
            $label = if ("$guidId" -match '\S') { Resolve-DirectoryObjectLabel -Id $guidId } else { 'Unnamed member' }
        }
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
    if ("$($Policy.Name)" -eq 'Evaluation Policy' -or "$($Policy.Name)" -like 'Defender for Office 365 Evaluation*') {
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
        'Default'             { 'Always on' }
        'Built-in protection' { 'On (built-in protection)' }
        'Evaluation'          { if ($Policy.PSObject.Properties['Enabled'] -and -not (Test-IsTrue $Policy.Enabled)) { 'Off (Defender for Office 365 evaluation)' } else { 'On - audit only (Defender for Office 365 evaluation; detects, takes no action)' } }
        default { if ($scopeRule) { if ("$($scopeRule.State)" -eq 'Enabled') { 'On' } elseif ("$($scopeRule.State)" -eq 'Disabled') { 'Off' } else { "$($scopeRule.State)" } } elseif ($RuleLookupFailed -and -not $scopeRule) { 'Unknown (rule lookup failed)' } else { 'Not applied (no rule)' } }
    }
    $inc = @(@($scopeRule.($f[0])) + @($scopeRule.($f[1])) + @($scopeRule.($f[2])) | Where-Object { $_ })
    $exc = @(@($scopeRule.($f[3])) + @($scopeRule.($f[4])) + @($scopeRule.($f[5])) | Where-Object { $_ })
    $applies = if ($kind -eq 'Evaluation') { 'Evaluation scope (see Defender for Office 365 evaluation settings)' }
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
    $sortBucket = switch ($kind) { 'Strict preset' { 0 } 'Standard preset' { 1 } 'Evaluation' { 2 } 'Built-in protection' { 4 } 'Default' { 5 } default { 3 } }
    $sortPrio = if ($sortBucket -eq 3 -and $scopeRule -and "$($scopeRule.Priority)" -match '^\d+$') { [int]"$($scopeRule.Priority)" } else { 99999 }
    [PSCustomObject]@{
        Kind = $kind; Status = $status; Priority = $(if ($kind -in 'Default','Built-in protection') { 'Lowest' } elseif ($kind -eq 'Evaluation') { 'Not applicable' } elseif ($kind -eq 'Strict preset') { 'Preset (evaluated first)' } elseif ($kind -eq 'Standard preset') { 'Preset (after Strict)' } elseif ($scopeRule) { $scopeRule.Priority } else { $null }); AppliesTo = $applies
        IncludedUsers = $scopeRule.($f[0]); IncludedGroups = $scopeRule.($f[1]); IncludedDomains = $scopeRule.($f[2])
        ExcludedUsers = $scopeRule.($f[3]); ExcludedGroups = $scopeRule.($f[4]); ExcludedDomains = $scopeRule.($f[5])
        SortKey = ($sortBucket * 100000 + $sortPrio)
    }
}

function Invoke-Section {
    param(
        [string]$Title,
        [int]$Level = 1,
        [scriptblock]$Body
    )
    Add-Line (('#' * $Level) + " " + $Title)
    Add-Line
    if ($Level -eq 1) { Write-Host "  $Title..." -ForegroundColor DarkGray }
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
        if ($server -and [int]$script:DnsBadServers[$server] -ge 2) { continue }
        $params = @{ Name = $Name; Type = $Type; DnsOnly = $true; QuickTimeout = $true; ErrorAction = "Stop" }
        if ($server) { $params.Server = $server }
        try {
            return @(DnsClient\Resolve-DnsName @params | Where-Object { $_.Section -eq 'Answer' -and "$($_.Type)" -eq $Type })
        }
        catch {
            # NXDOMAIN is an expected "no record" result; anything else is logged so empty DNS cells can be explained.
            if ("$($_.FullyQualifiedErrorId)" -match '^DNS_ERROR_RCODE_NAME_ERROR|^DNS_INFO_NO_RECORDS' -or $_.Exception.NativeErrorCode -in 9003, 9501 -or $_.Exception.Message -match 'does not exist') { return @() }
            $lastError = $_
            if ($server -and -not $DnsServer) { $script:DnsBadServers[$server] = [int]$script:DnsBadServers[$server] + 1 }
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

function Get-SharedMailboxes {
    if ($script:MailboxesError) { throw $script:MailboxesError }
    if ($null -eq $script:Mailboxes) {
        $props = @('RecipientTypeDetails', 'UserPrincipalName', 'PrimarySmtpAddress', 'DisplayName',
                   'LitigationHoldEnabled', 'ArchiveStatus', 'AutoExpandingArchiveEnabled', 'RetentionHoldEnabled',
                   'AuditEnabled', 'ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward',
                   'RoleAssignmentPolicy', 'RetentionPolicy', 'WhenMailboxCreated', 'IsInactiveMailbox', 'ProhibitSendReceiveQuota')
        try { $script:Mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $props) }
        catch { $script:MailboxesError = $_; throw }
    }
    return $script:Mailboxes
}

function Get-SharedCasMailboxes {
    if ($script:CasMailboxesError) { throw $script:CasMailboxesError }
    if ($null -eq $script:CasMailboxes) {
        $props = @('PopEnabled', 'ImapEnabled', 'EwsEnabled', 'ActiveSyncEnabled', 'MAPIEnabled', 'OWAEnabled',
                   'SmtpClientAuthenticationDisabled', 'OwaMailboxPolicy', 'ActiveSyncMailboxPolicy')
        try { $script:CasMailboxes = @(Get-EXOCASMailbox -ResultSize Unlimited -Properties $props) }
        catch { $script:CasMailboxesError = $_; throw }
    }
    return $script:CasMailboxes
}

function Get-SharedDistributionGroups {
    if ($script:DistGroupsError) { throw $script:DistGroupsError }
    if ($null -eq $script:DistGroups) {
        try { $script:DistGroups = @(Get-DistributionGroup -ResultSize Unlimited) } catch { $script:DistGroupsError = $_; throw }
    }
    return $script:DistGroups
}

function Get-SharedDynamicGroups {
    if ($script:DynGroupsError) { throw $script:DynGroupsError }
    if ($null -eq $script:DynGroups) {
        try { $script:DynGroups = @(Get-DynamicDistributionGroup -ResultSize Unlimited) } catch { $script:DynGroupsError = $_; throw }
    }
    return $script:DynGroups
}

function Get-SharedM365Groups {
    if ($script:M365GroupsError) { throw $script:M365GroupsError }
    if ($null -eq $script:M365Groups) {
        try { $script:M365Groups = @(Get-UnifiedGroup -ResultSize Unlimited) } catch { $script:M365GroupsError = $_; throw }
    }
    return $script:M365Groups
}

function Get-SharedInboundConnectors {
    if ($script:InboundConnectorsError) { throw $script:InboundConnectorsError }
    if ($null -eq $script:InboundConnectors) {
        try { $script:InboundConnectors = @(Get-InboundConnector) } catch { $script:InboundConnectorsError = $_; throw }
    }
    return $script:InboundConnectors
}

function Get-SharedOutboundConnectors {
    if ($script:OutboundConnectorsError) { throw $script:OutboundConnectorsError }
    if ($null -eq $script:OutboundConnectors) {
        try { $script:OutboundConnectors = @(Get-OutboundConnector) } catch { $script:OutboundConnectorsError = $_; throw }
    }
    return $script:OutboundConnectors
}

function Get-SharedTransportRules {
    if ($script:TransportRulesError) { throw $script:TransportRulesError }
    if ($null -eq $script:TransportRules) {
        try { $script:TransportRules = @(Get-TransportRule) } catch { $script:TransportRulesError = $_; throw }
    }
    return $script:TransportRules
}

function Get-SharedTransportConfig {
    if ($script:TransportConfigError) { throw $script:TransportConfigError }
    if ($null -eq $script:TransportConfig) {
        try { $script:TransportConfig = @(Get-TransportConfig) } catch { $script:TransportConfigError = $_; throw }
    }
    return $script:TransportConfig
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

# ============================================================
# Markdown formatting helpers (Word-friendly output)
# ============================================================

function Test-IsTrue {
    param($Value)
    if ($Value -is [bool]) { return $Value }
    return ("$Value" -eq 'True')
}

function Get-NonEmptyCount { param($Value) return @($Value | Where-Object { "$_" -match '\S' }).Count }

function Get-FirstProp {
    param($Obj, [string[]]$Names)
    foreach ($n in $Names) { if ($Obj.PSObject.Properties.Name -contains $n -and "$($Obj.$n)" -match '\S') { return $Obj.$n } }
    return $null
}

function ConvertTo-IntegratedAppRows {
    param([object[]]$List, [hashtable]$Details)
    return @($List | ForEach-Object {
        $item = $_
        $detail = $item
        if ($item.ProductId -and $Details.ContainsKey("$($item.ProductId)")) { $detail = $Details["$($item.ProductId)"] }
        $name = Get-FirstProp $detail @('DisplayName','Name','Title')
        if (-not $name) { $name = "$($detail.ProductId)" }
        $enabled = Get-FirstProp $detail @('Enabled','IsEnabled')
        $assigned = 'Not assigned'
        $everyone = Get-FirstProp $detail @('AssignToEveryone','AssignedToEveryone','IsAssignedToEveryone')
        if (Test-IsTrue $everyone) { $assigned = 'Everyone' }
        else {
            $assignees = @('Members','AssignedUsers','AssignedGroups','Users','Groups','Assignments' | ForEach-Object { @($detail.$_ | Where-Object { "$_" }) } | ForEach-Object { $_ } | Where-Object { $_ })
            if ($assignees.Count) { $assigned = "$($assignees.Count) user(s)/group(s): " + ((@($assignees | Select-Object -First 3)) -join ', ') + $(if ($assignees.Count -gt 3) { ' …' } else { '' }) }
        }
        $o = [ordered]@{ Name = $name; Enabled = $enabled; AssignedTo = $assigned; ProductId = $detail.ProductId }
        foreach ($p in $detail.PSObject.Properties.Name) { if (-not $o.Contains($p)) { $o[$p] = $detail.$p } }
        [PSCustomObject]$o
    })
}

function Format-Value {
    param($Value, [string]$CsvFile)
    if ($null -eq $Value) { return "Not configured" }
    if ($Value -is [bool]) { return $(if ($Value) { 'Yes' } else { 'No' }) }
    if ($Value -is [string] -and $Value.Trim() -match '^(True|False)$') { return $(if (Test-IsTrue $Value) { 'Yes' } else { 'No' }) }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $ts = $null
    if ($Value -is [timespan]) { $ts = $Value }
    elseif ($Value -is [string] -and $Value -match '^(\d+\.\d{2}:\d{2}:\d{2}|\d{2}:\d{2}:\d{2})$') { $ts = [timespan]::Parse($Value) }
    if ($null -ne $ts) {
        if ($ts -eq [timespan]::Zero) { return '0' }
        if ($ts.TotalDays -eq [math]::Floor($ts.TotalDays)) { return ("$([int]$ts.TotalDays) day" + $(if ([int]$ts.TotalDays -eq 1) { '' } else { 's' })) }
        if ($ts.TotalHours -eq [math]::Floor($ts.TotalHours)) { return ("$([int]$ts.TotalHours) hour" + $(if ([int]$ts.TotalHours -eq 1) { '' } else { 's' })) }
        if ($ts.TotalMinutes -eq [math]::Floor($ts.TotalMinutes)) { return ("$([int]$ts.TotalMinutes) minutes") }
        return "$([math]::Round($ts.TotalSeconds, 0)) seconds"
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = @($Value | ForEach-Object { "$_" })
        if ($items.Count -eq 0) { return "Not configured" }
        if ($items.Count -gt $MaxListItems) {
            $shown = @($items | Select-Object -First $MaxListItems)
            $more = $items.Count - $MaxListItems
            $suffix = if ($CsvFile) { "see CSV ``$CsvFile``" } else { "see CSV export" }
            return ((($shown | ForEach-Object { $_ -replace '\|', '\|' -replace '<', '\<' -replace "`r?`n", ' ' }) -join ", ") + " … (+$more more – $suffix)")
        }
        return (($items | ForEach-Object { $_ -replace '\|', '\|' -replace '<', '\<' -replace "`r?`n", ' ' }) -join ", ")
    }
    $text = "$Value" -replace "`r`n", "; " -replace "`n", "; " -replace "`r", "; " -replace "`t", ' '
    $text = $text -replace ';\s*$', '' -replace '\s+$', ''
    if ($text -match '^(.+?)\s*\([\d,.\s]+bytes\)$') { $text = $Matches[1] }
    if ($text -notmatch '\S') { return "Not configured" }
    return ($text -replace '\|', '\|' -replace '<', '\<')
}

function ConvertTo-MdHeading { param($Text) ("$Text" -replace '<', '\<') }

function Out-Card {
    # Renders an ordered label -> raw value map as a vertical Setting|Value table.
    param([System.Collections.IDictionary]$Settings)
    Add-Line "| Setting | Value |"
    Add-Line "|---|---|"
    foreach ($key in $Settings.Keys) {
        $entry = $Settings[$key]
        if ($entry -is [System.Collections.IDictionary] -and $entry.Contains('Value')) {
            Add-Line "| $(($key -replace '\|', '\|')) | $(Format-Value $entry.Value -CsvFile $entry.CsvFile) |"
        }
        else {
            Add-Line "| $(($key -replace '\|', '\|')) | $(Format-Value $entry) |"
        }
    }
    Add-Line
}

function Register-Csv {
    param([string]$Name, [object[]]$Rows, [string]$Description, [switch]$Failed, [string]$Note)
    if ("$Note" -match '\S') { $Description = "$Description ($Note)" }
    if ($Failed) {
        $script:CsvIndex[$Name] = [PSCustomObject]@{ File = 'not collected (lookup failed)'; Contents = $Description; Rows = 0 }
        return
    }
    $fileName = if (@($Rows).Count -gt 0) { "$Name.csv" } else { "none" }
    $script:CsvIndex[$Name] = [PSCustomObject]@{ File = $fileName; Contents = $Description; Rows = @($Rows).Count }
    if (@($Rows).Count -gt 0) {
        $script:CsvData[$Name] = @($Rows)
    }
}

function ConvertTo-HeaderText {
    param([string]$Name)
    $map = @{
        AppliesTo = 'Applies to'; DisplayName = 'Display name'; PrimarySmtpAddress = 'Email address'
        RecipientFilter = 'Recipient filter'; RoleAssignmentPolicy = 'Role assignment policy'
        MailboxCount = 'Mailboxes'; EnableATPForSPOTeamsODB = 'ATP for SPO/OneDrive/Teams'
        EnableSafeDocs = 'Safe Documents'; ExpirationDate = 'Expires'; ScopeRestrictionType = 'Scope type'
        IsDefault = 'Default'; SimulationUrl = 'Simulation URL'; ConnectorType = 'Connector type'
        UseMXRecord = 'Use MX record'; JournalEmailAddress = 'Journal email address'
        AutoForwardEnabled = 'Auto-forward'; AutoReplyEnabled = 'Auto-reply'
        AllowedOOFType = 'Automatic replies'; RecipientMessageAccess = 'Recipient message access'
        QuarantineNotification = 'Quarantine notification'; RetentionPeriod = 'Retention period'
        RetentionAction = 'Action'; IsRecordLabel = 'Record label'; ExchangeLocation = 'Exchange locations'
        AutomateProcessing = 'Automatic processing'; BookingWindowInDays = 'Booking window (days)'
        MaximumDurationInMinutes = 'Max duration (minutes)'; MaxDuration = 'Max duration (minutes)'; DefaultStateForUser = 'Default state for user'
        ProhibitSendReceiveQuota = 'Mailbox quota'
        ProvidedTo = 'Provided to'; Policy = 'Policy'
        ForwardTo = 'Forwards to'; KeepCopy = 'Keeps a copy'
        AssignedTo = 'Assigned to'; ProductId = 'Product ID'
    }
    if ($map.ContainsKey($Name)) { return $map[$Name] }
    $words = [regex]::Matches($Name, '[A-Z]+(?![a-z])|[A-Z][a-z0-9]*|[a-z0-9]+')
    if ($words.Count -eq 0) { return $Name }
    $out = @()
    for ($i = 0; $i -lt $words.Count; $i++) {
        $w = $words[$i].Value
        if ($i -eq 0) { $out += $w.Substring(0,1).ToUpper() + $w.Substring(1) }
        elseif ($w -cmatch '^[A-Z]{2,}') { $out += $w }
        else { $out += $w.ToLower() }
    }
    return ($out -join ' ')
}

function Format-ThreatAction {
    param($Value)
    $map = @{
        MoveToJmf     = "Move message to the recipients' Junk Email folders"
        Quarantine    = 'Quarantine the message'
        NoAction      = "Don't apply any action"
        Delete        = 'Delete the message'
        AddXHeader    = 'Add X-header'
        ModifySubject = 'Prepend subject line with text'
        Redirect      = 'Redirect message to email address'
        Reject        = 'Reject the message'
        BccMessage    = 'Add recipients to the Bcc box'
    }
    if ($map.ContainsKey("$Value")) { return $map["$Value"] }
    return $Value
}

function Format-ConnectorDomain {
    param($Domains)
    @($Domains | ForEach-Object {
        $d = "$_" -replace '(?i)^smtp:', '' -replace ';\d+$', ''
        if ($d -eq '*') { 'All domains (*)' } else { $d }
    })
}

function Format-SharingDomains {
    param($Domains)
    $actionMap = @{
        CalendarSharingFreeBusySimple   = 'Calendar: free/busy only'
        CalendarSharingFreeBusyDetail   = 'Calendar: limited details'
        CalendarSharingFreeBusyReviewer = 'Calendar: all information'
        ContactsSharing                 = 'Contacts'
    }
    $out = @()
    foreach ($e in @($Domains)) {
        $entry = "$e"
        $parts = $entry -split ':', 2
        $dom = $parts[0]
        $acts = if ($parts.Count -gt 1) { $parts[1] } else { '' }
        $domText = switch ($dom) {
            '*'         { 'All domains' }
            'Anonymous' { 'Anonymous (published calendars)' }
            default     { $dom }
        }
        $actText = @($acts -split ',' | Where-Object { $_ } | ForEach-Object {
            $a = "$_".Trim()
            if ($actionMap.ContainsKey($a)) { $actionMap[$a] } else { $a }
        }) -join ', '
        if ($actText) { $out += "$domText`: $actText" } else { $out += $domText }
    }
    return ($out -join '; ')
}

function Out-Overview {
    # Capped <=4-column table; registers the full row set for CSV export.
    param(
        [string]$CsvName,
        [string]$Description,
        [object[]]$Rows,
        [string[]]$Columns,
        [switch]$SkipCsv,
        [switch]$Failed,
        [string]$Note
    )
    if ($Columns.Count -gt 4) { throw "Out-Overview '$CsvName' has more than 4 columns" }
    if (-not $SkipCsv) { Register-Csv -Name $CsvName -Rows $Rows -Description $Description -Failed:$Failed -Note $Note }
    if ($Failed) {
        Add-Line "_Not available - lookup failed (see collection log)._"
        Add-Line
        return
    }
    $rows = @($Rows)
    if ($rows.Count -eq 0) {
        Add-Line "_None found._"
        Add-Line
        return
    }
    $shown = $rows
    $truncated = $false
    if ($rows.Count -gt $MaxTableRows) {
        $shown = @($rows | Select-Object -First $MaxTableRows)
        $truncated = $true
    }
    Add-Line ("| " + (@($Columns | ForEach-Object { ConvertTo-HeaderText $_ }) -join " | ") + " |")
    Add-Line ("|" + (($Columns | ForEach-Object { "---" }) -join "|") + "|")
    foreach ($row in $shown) {
        $cells = foreach ($col in $Columns) { Format-Value $row.$col }
        Add-Line ("| " + ($cells -join " | ") + " |")
    }
    if ($truncated) {
        Add-Line
        Add-Line "_Showing $($shown.Count) of $($rows.Count) rows – see CSV ``$CsvName.csv`` for the full list._"
    }
    Add-Line
}

# ============================================================
# Document sections
# ============================================================

function Write-DocumentInfoSection {
    Invoke-Section -Title 'Document information' -Body {
        Add-Line "This document describes the Exchange Online configuration of $CustomerName as collected on $($DocumentDate.ToString('yyyy-MM-dd')). It is a read-only documentation generated by Invoke-ExchangeOnlineDocumentation.ps1."
        Add-Line
        $purviewState = if (-not $IncludePurview) { 'Not collected (run with -IncludePurview)' } elseif ($script:PurviewConnected) { 'Collected' } else { "Failed ($script:PurviewError)" }
        Out-Card -Settings ([ordered]@{
            'Customer'              = $CustomerName
            'Tenant display name'   = $script:TenantName
            'Tenant ID'             = $script:TenantId
            'Initial (MOERA) domain' = $script:InitialDomain
            'Documentation date'    = $DocumentDate.ToString('yyyy-MM-dd')
            'Author'                = $Author
            'Collected by'          = $script:CollectedBy
            'EXO module version'    = $script:ExoModuleVersion
            'Script version'        = $script:ScriptVersion
            'Purview data'          = $purviewState
            'Integrated apps data'  = $(if (-not $IncludeIntegratedApps) { 'Not collected (run with -IncludeIntegratedApps)' } elseif ($script:IntegratedAppsConnected) { 'Collected' } else { "Failed ($script:IntegratedAppsError)" })
        })
    }
}

function Write-TenantOverviewSection {
    Invoke-Section -Title 'Tenant overview' -Body {
        Add-Line "A quick overview of the size and shape of the tenant: how many mailboxes, groups, domains and mail flow objects exist."
        Add-Line
        $mbx = @(); $mbxFailed = $false
        try { $mbx = @(Get-SharedMailboxes) } catch { $mbxFailed = $true; Add-LogEntry -Section 'Mailbox count' -ErrorRecord $_ }
        $dg = @(); $dyn = @(); $m365 = @(); $dgFailed = $false; $dynFailed = $false; $m365Failed = $false
        try { $dg = @(Get-SharedDistributionGroups) } catch { $dgFailed = $true; Add-LogEntry -Section 'Group count' -ErrorRecord $_ }
        try { $dyn = @(Get-SharedDynamicGroups) } catch { $dynFailed = $true; Add-LogEntry -Section 'Dynamic group count' -ErrorRecord $_ }
        try { $m365 = @(Get-SharedM365Groups) } catch { $m365Failed = $true; Add-LogEntry -Section 'M365 group count' -ErrorRecord $_ }
        $inC = @(); $outC = @(); $tr = @(); $inCFailed = $false; $outCFailed = $false; $trFailed = $false
        try { $inC = @(Get-SharedInboundConnectors) } catch { $inCFailed = $true; Add-LogEntry -Section 'Inbound connector count' -ErrorRecord $_ }
        try { $outC = @(Get-SharedOutboundConnectors) } catch { $outCFailed = $true; Add-LogEntry -Section 'Outbound connector count' -ErrorRecord $_ }
        try { $tr = @(Get-SharedTransportRules) } catch { $trFailed = $true; Add-LogEntry -Section 'Transport rule count' -ErrorRecord $_ }
        $doms = @(); $domsFailed = $false
        try { $doms = @(if ($script:AcceptedDomains) { $script:AcceptedDomains } else { Get-AcceptedDomain }) }
        catch { $domsFailed = $true; Add-LogEntry -Section 'Accepted domain count' -ErrorRecord $_ }
        $mbxC = $(if ($mbxFailed) { $script:LookupFailedText } else { $null })
        Out-Card -Settings ([ordered]@{
            'Mailboxes (total)'            = $(if ($mbxFailed) { $mbxC } else { $mbx.Count })
            '  User mailboxes'             = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }).Count })
            '  Shared mailboxes'           = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count })
            '  Room mailboxes'             = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox' }).Count })
            '  Equipment mailboxes'        = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'EquipmentMailbox' }).Count })
            '  Scheduling (Bookings) mailboxes' = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'SchedulingMailbox' }).Count })
            '  Discovery mailboxes'        = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'DiscoveryMailbox' }).Count })
            '  Other mailbox types'        = $(if ($mbxFailed) { $mbxC } else { @($mbx | Where-Object { $_.RecipientTypeDetails -notin 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','SchedulingMailbox','DiscoveryMailbox' }).Count })
            'Distribution groups'          = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup' }).Count })
            'Mail-enabled security groups' = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup' }).Count })
            'Room lists'                   = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' }).Count })
            'Dynamic distribution groups'  = $(if ($dynFailed) { $script:LookupFailedText } else { $dyn.Count })
            'Microsoft 365 groups'         = $(if ($m365Failed) { $script:LookupFailedText } else { $m365.Count })
            'Accepted domains'             = $(if ($domsFailed) { $script:LookupFailedText } else { $doms.Count })
            'Inbound connectors'           = $(if ($inCFailed) { $script:LookupFailedText } else { $inC.Count })
            'Outbound connectors'          = $(if ($outCFailed) { $script:LookupFailedText } else { $outC.Count })
            'Transport rules'              = $(if ($trFailed) { $script:LookupFailedText } else { $tr.Count })
        })
    }
}

function Write-OrgSettingsSection {
    Invoke-Section -Title 'Organization settings' -Body {
        Add-Line "Tenant-wide Exchange Online settings: general organisation flags, transport limits, auditing and external sender tagging."
        Add-Line

        Add-Line "## General"
        Add-Line
        $orgSettings = @(
            @{ Label = 'Organization name'; Prop = 'DisplayName' },
            @{ Label = 'Modern authentication enabled'; Prop = 'OAuth2ClientProfileEnabled' },
            @{ Label = 'Customer Lockbox enabled'; Prop = 'CustomerLockBoxEnabled' },
            @{ Label = 'MailTips enabled'; Prop = 'MailTipsAllTipsEnabled' },
            @{ Label = 'MailTip for external recipients'; Prop = 'MailTipsExternalRecipientsTipsEnabled' },
            @{ Label = 'Large audience MailTip threshold'; Prop = 'MailTipsLargeAudienceThreshold' },
            @{ Label = 'EWS allowed applications'; Prop = 'EwsAllowList' },
            @{ Label = 'Focused Inbox on'; Prop = 'FocusedInboxOn' },
            @{ Label = 'Default authentication policy'; Prop = 'DefaultAuthenticationPolicy' },
            @{ Label = 'Activity-based sign-out enabled'; Prop = 'ActivityBasedAuthenticationTimeoutEnabled' },
            @{ Label = 'Send from alias enabled'; Prop = 'SendFromAliasEnabled' },
            @{ Label = 'Auto-expanding archiving (org-wide)'; Prop = 'AutoExpandingArchiveEnabled' }
        )
        $org = $null; $orgFailed = $false
        try { $org = if ($script:OrgConfig) { $script:OrgConfig } else { Get-OrganizationConfig } }
        catch { $orgFailed = $true; Add-LogEntry -Section 'OrganizationConfig' -ErrorRecord $_ }
        if ($org) {
            $card = [ordered]@{}
            $card[$orgSettings[0].Label] = $org.DisplayName
            $card[$orgSettings[1].Label] = $org.OAuth2ClientProfileEnabled
            $card['Mailbox auditing on by default'] = $(if ($null -eq $org.AuditDisabled) { $null } else { -not (Test-IsTrue $org.AuditDisabled) })
            $card[$orgSettings[2].Label] = $org.CustomerLockBoxEnabled
            $card[$orgSettings[3].Label] = $org.MailTipsAllTipsEnabled
            $card[$orgSettings[4].Label] = $org.MailTipsExternalRecipientsTipsEnabled
            $card[$orgSettings[5].Label] = $org.MailTipsLargeAudienceThreshold
            $card['EWS enabled (org-wide)'] = $(if ($null -eq $org.EwsEnabled) { 'Not set (enabled)' } else { $org.EwsEnabled })
            $card['EWS application access policy'] = switch ("$($org.EwsApplicationAccessPolicy)") {
                ''                 { 'Not set (all applications allowed)' }
                'EnforceAllowList' { 'Allow list enforced (only listed applications allowed)' }
                'EnforceBlockList' { 'Block list enforced (listed applications blocked)' }
                default            { $org.EwsApplicationAccessPolicy }
            }
            $card[$orgSettings[6].Label] = $org.EwsAllowList
            $card['EWS blocked applications'] = $org.EwsBlockList
            $card[$orgSettings[7].Label] = $org.FocusedInboxOn
            $card[$orgSettings[8].Label] = $org.DefaultAuthenticationPolicy
            $card[$orgSettings[9].Label] = $org.ActivityBasedAuthenticationTimeoutEnabled
            $card[$orgSettings[10].Label] = $org.SendFromAliasEnabled
            $card[$orgSettings[11].Label] = $org.AutoExpandingArchiveEnabled
            Out-Card -Settings $card
            Register-Csv -Name 'OrganizationConfig' -Description 'Organization-wide settings' -Rows @($org) -Failed:$orgFailed
        }
        elseif ($orgFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not available._"; Add-Line }

        Add-Line "## Transport"
        Add-Line
        $tcSettings = @(
            @{ Label = 'SMTP AUTH disabled (org-wide)'; Prop = 'SmtpClientAuthenticationDisabled' },
            @{ Label = 'Allow legacy TLS clients'; Prop = 'AllowLegacyTLSClients' },
            @{ Label = 'Max send size'; Prop = 'MaxSendSize' },
            @{ Label = 'Max receive size'; Prop = 'MaxReceiveSize' },
            @{ Label = 'Max recipients per message'; Prop = 'MaxRecipientEnvelopeLimit' },
            @{ Label = 'External postmaster address'; Prop = 'ExternalPostmasterAddress' },
            @{ Label = 'Journaling NDR address'; Prop = 'JournalingReportNdrTo' },
            @{ Label = 'Reply-all storm protection'; Prop = 'ReplyAllStormProtectionEnabled' },
            @{ Label = 'Delay notifications to external senders'; Prop = 'ExternalDelayDsnEnabled' },
            @{ Label = 'Message expiration timeout'; Prop = 'MessageExpiration' }
        )
        $tc = @(); $tcFailed = $false
        try { $tc = @(Get-SharedTransportConfig) } catch { $tcFailed = $true; Add-LogEntry -Section 'TransportConfig' -ErrorRecord $_ }
        if ($tc.Count -gt 0) {
            $card = [ordered]@{}
            $tcObj = $tc[0]
            foreach ($s in $tcSettings) { $card[$s.Label] = $tcObj.($s.Prop) }
            if ("$($card['Journaling NDR address'])" -in '','<>') { $card['Journaling NDR address'] = $null }
            Out-Card -Settings $card
            Register-Csv -Name 'TransportConfig' -Description 'Transport (org-wide mail flow) settings' -Rows $tc -Failed:$tcFailed
        }
        elseif ($tcFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not available._"; Add-Line }

        Add-Line "## Auditing"
        Add-Line
        $aal = @(); $aalFailed = $false
        try { $aal = @(Get-AdminAuditLogConfig) } catch { $aalFailed = $true; Add-LogEntry -Section 'AdminAuditLogConfig' -ErrorRecord $_ }
        if ($aal.Count -gt 0) {
            $card = [ordered]@{}
            $aalObj = $aal[0]
            $card['Unified audit log search enabled'] = $aalObj.UnifiedAuditLogIngestionEnabled
            $card['Admin audit log enabled'] = $aalObj.AdminAuditLogEnabled
            $card['Admin audit log retention'] = $aalObj.AdminAuditLogAgeLimit
            $mbxAudit = @(); $mbxAuditFailed = $false
            try { $mbxAudit = @(Get-SharedMailboxes) } catch { $mbxAuditFailed = $true }
            $card['Mailboxes with audit enabled'] = $(if ($mbxAuditFailed) { $script:LookupFailedText } else { @($mbxAudit | Where-Object { Test-IsTrue $_.AuditEnabled }).Count })
            $card['Mailboxes with audit disabled'] = $(if ($mbxAuditFailed) { $script:LookupFailedText } else { @($mbxAudit | Where-Object { "$($_.AuditEnabled)" -eq 'False' }).Count })
            $auditUnset = @($mbxAudit | Where-Object { $null -eq $_.AuditEnabled -or "$($_.AuditEnabled)" -eq '' }).Count
            if ($auditUnset -gt 0) { $card['Mailboxes with audit not reported'] = $auditUnset }
            Out-Card -Settings $card
            Register-Csv -Name 'AdminAuditLogConfig' -Description 'Audit log configuration' -Rows $aal -Failed:$aalFailed
        }
        elseif ($aalFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not available._"; Add-Line }

        Add-Line "## External sender tagging"
        Add-Line
        $ext = @(); $extFailed = $false
        try { $ext = @(Get-ExternalInOutlook) } catch { $extFailed = $true; Add-LogEntry -Section 'ExternalInOutlook' -ErrorRecord $_ }
        $extRows = @($ext | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; Enabled = $_.Enabled; AllowList = $_.AllowList }
        })
        if ($extRows.Count -gt 0) {
            $card = [ordered]@{}
            foreach ($e in $extRows) {
                $card['Enabled'] = $e.Enabled
                $card['Allow list'] = @{ Value = $e.AllowList; CsvFile = 'ExternalInOutlook.csv' }
            }
            Out-Card -Settings $card
        }
        elseif ($extFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not configured._"; Add-Line }
        Register-Csv -Name 'ExternalInOutlook' -Description 'External sender tagging' -Rows $extRows -Failed:$extFailed
    }
}

function Write-DomainsSection {
    Invoke-Section -Title 'Domains and email authentication' -Body {
        Add-Line "Accepted domains in Exchange Online and the state of email authentication (MX, SPF, DKIM, DMARC, MTA-STS, TLS-RPT) per domain, checked against public DNS."
        Add-Line

        $domains = @(); $domsFailed = $false
        try { $domains = @(if ($script:AcceptedDomains) { $script:AcceptedDomains } else { Get-AcceptedDomain }) }
        catch { $domsFailed = $true; Add-LogEntry -Section 'AcceptedDomain' -ErrorRecord $_ }
        $domRows = @($domains | ForEach-Object {
            [PSCustomObject]@{ Domain = "$($_.DomainName)".Trim(); Type = $_.DomainType; Default = $_.Default; Initial = $_.InitialDomain }
        })
        Add-Line "## Accepted domains"
        Add-Line
        Out-Overview -CsvName 'AcceptedDomains' -Description 'Accepted domains' -Rows $domRows -Columns @('Domain','Type','Default','Initial') -Failed:$domsFailed

        $dkim = @(); $dkimFailed = $false
        try { $dkim = @(Get-DkimSigningConfig) } catch { $dkimFailed = $true; Add-LogEntry -Section 'DKIM config' -ErrorRecord $_ }
        $script:DkimConfig = $dkim

        Add-Line "## Per-domain configuration"
        Add-Line
        $dnsRows = @()
        foreach ($d in $domains) {
            $name = "$($d.DomainName)".Trim()
            if ($name -like "*.onmicrosoft.com") { continue }
            $mx = @(Resolve-DnsSafe -Name $name -Type MX | ForEach-Object { $_.NameExchange })
            $mxExo = @($mx | Where-Object { "$_".TrimEnd('.') -match '(?i)\.mail\.protection\.outlook\.com$|\.mx\.microsoft$' }).Count -gt 0
            $spf = @(Get-TxtRecords -Name $name | Where-Object { $_ -like "v=spf1*" })
            $spfText = ($spf -join '; ') + $(if ($spf.Count -gt 1) { ' (multiple SPF records - invalid)' } else { '' })
            $dmarcTxt = @(Get-TxtRecords -Name "_dmarc.$name" | Where-Object { $_ -like "v=DMARC1*" })
            $mtaSts = @(Get-TxtRecords -Name "_mta-sts.$name" | Where-Object { $_ -like "v=STSv1*" })
            $tlsRpt = @(Get-TxtRecords -Name "_smtp._tls.$name" | Where-Object { $_ -like "v=TLSRPTv1*" })

            $dmarcParts = @{}
            if ($dmarcTxt.Count -gt 0) {
                foreach ($kv in ("$($dmarcTxt[0])" -split ';')) {
                    if ($kv -match '^\s*(\w+)\s*=\s*(.+?)\s*$') { $dmarcParts[$Matches[1]] = $Matches[2] }
                }
            }
            $dk = @($dkim | Where-Object { "$($_.Domain)" -eq $name }) | Select-Object -First 1
            $dkimDns = @{}
            foreach ($sel in 1, 2) {
                $host_ = @(Resolve-DnsSafe -Name "selector$sel._domainkey.$name" -Type CNAME | ForEach-Object { $_.NameHost }) | Select-Object -First 1
                $expected = if ($sel -eq 1) { $dk.Selector1CNAME } else { $dk.Selector2CNAME }
                $dkimDns[$sel] = if (-not $host_) { 'Not published' }
                    elseif (-not $dk -or "$expected".Trim() -eq '') { "Published (target: $host_)" }
                    elseif (("$host_".TrimEnd('.')) -ieq ("$expected".TrimEnd('.'))) { 'Published (matches)' }
                    else { "Published (different target: $host_)" }
            }
            $dnsRows += [PSCustomObject]@{
                Domain = $name; MX = ($mx -join '; '); MXPointsToEXO = $mxExo
                SPF = $spfText; DMARC_p = $dmarcParts['p']; DMARC_sp = $dmarcParts['sp']
                DMARC_pct = $dmarcParts['pct']; DMARC_rua = $dmarcParts['rua']; DMARC_ruf = $dmarcParts['ruf']
                DMARC_RecordCount = $dmarcTxt.Count
                'MTA-STS' = ($mtaSts -join '; '); 'TLS-RPT' = ($tlsRpt -join '; ')
                DkimSelector1Dns = $dkimDns[1]; DkimSelector2Dns = $dkimDns[2]
            }

            Add-Line "### $(ConvertTo-MdHeading $name)"
            Add-Line
            $domCard = [ordered]@{
                'Domain type'        = $d.DomainType
                'MX'                 = @{ Value = $mx; CsvFile = 'DomainDns.csv' }
                'MX points to EXO'   = $mxExo
                'SPF'                = $spfText
                'DKIM enabled'       = if ($dkimFailed) { $script:LookupFailedText } elseif ($dk) { $dk.Enabled } else { $null }
                'DKIM selector 1'    = if ($dkimFailed) { $script:LookupFailedText } elseif ($dk) { $dk.Selector1CNAME } else { $null }
                'DKIM selector 2'    = if ($dkimFailed) { $script:LookupFailedText } elseif ($dk) { $dk.Selector2CNAME } else { $null }
                'DKIM selector 1 CNAME in DNS' = $dkimDns[1]
                'DKIM selector 2 CNAME in DNS' = $dkimDns[2]
                'DMARC policy'       = $dmarcParts['p']
                'DMARC pct'          = $dmarcParts['pct']
                'DMARC rua'          = $dmarcParts['rua']
                'MTA-STS'            = ($mtaSts -join '; ')
                'TLS-RPT'            = ($tlsRpt -join '; ')
            }
            if ($dmarcTxt.Count -gt 1) { $domCard['DMARC records'] = "$($dmarcTxt.Count) records found - invalid (only one allowed)" }
            Out-Card -Settings $domCard
        }
        Register-Csv -Name 'DomainDns' -Description 'DNS authentication records per accepted domain' -Rows $dnsRows -Failed:$domsFailed

        Add-Line "## DKIM signing configuration"
        Add-Line
        $dkimRows = @($dkim | ForEach-Object {
            [PSCustomObject]@{
                Domain = $_.Domain; Enabled = $_.Enabled; Selector1CNAME = $_.Selector1CNAME
                Selector2CNAME = $_.Selector2CNAME; KeySize = $_.KeySize; LastChecked = $_.LastChecked
                RotateOnDate = $_.RotateOnDate; SelectorBeforeRotateOnDate = $_.SelectorBeforeRotateOnDate
            }
        })
        Out-Overview -CsvName 'DkimConfig' -Description 'DKIM signing configuration per domain' -Rows $dkimRows -Columns @('Domain','Enabled','Selector1CNAME','Selector2CNAME') -Failed:$dkimFailed

        Add-Line "## MOERA (onmicrosoft.com) DKIM/DMARC"
        Add-Line
        if ($script:InitialDomain) {
            $moeraDmarc = @(Get-TxtRecords -Name "_dmarc.$($script:InitialDomain)" | Where-Object { $_ -like "v=DMARC1*" })
            $moeraDkim = @($dkim | Where-Object { $_.Domain -eq $script:InitialDomain }) | Select-Object -First 1
            Out-Card -Settings ([ordered]@{
                'Domain'      = $script:InitialDomain
                'DMARC'       = ($moeraDmarc -join '; ')
                'DKIM enabled' = if ($dkimFailed) { $script:LookupFailedText } elseif ($moeraDkim) { $moeraDkim.Enabled } else { $null }
            })
            Register-Csv -Name 'MoeraAuth' -Description 'MOERA domain DKIM/DMARC' -Rows @(
                [PSCustomObject]@{ Domain = $script:InitialDomain; DMARC = ($moeraDmarc -join '; '); DKIMEnabled = if ($moeraDkim) { $moeraDkim.Enabled } else { $null } })
        }
        else {
            Add-Line "_Initial (MOERA) domain not identified._"
            Add-Line
            Register-Csv -Name 'MoeraAuth' -Description 'MOERA domain DKIM/DMARC' -Rows @()
        }

        Add-Line "## ARC trusted sealers"
        Add-Line
        $arc = @(); $arcFailed = $false
        try { $arc = @(Get-ArcConfig) } catch { $arcFailed = $true; Add-LogEntry -Section 'ARC config' -ErrorRecord $_ }
        $arcRows = @($arc | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; ArcTrustedSealers = $_.ArcTrustedSealers }
        })
        Out-Overview -CsvName 'ArcConfig' -Description 'ARC trusted sealers' -Rows $arcRows -Columns @('Identity','ArcTrustedSealers') -Failed:$arcFailed
    }
}

function Write-MailFlowSection {
    Invoke-Section -Title 'Mail flow' -Body {
        Add-Line "Connectors, transport rules, remote domains, journal rules and High Volume Email (HVE) accounts - i.e. how mail enters, leaves and is processed in the tenant."
        Add-Line

        Add-Line "## Inbound connectors"
        Add-Line
        $inProps = @('ConnectorType', 'Enabled', 'SenderDomains', 'SenderIPAddresses', 'RequireTls',
                     'TlsSenderCertificateName', 'RestrictDomainsToIPAddresses', 'RestrictDomainsToCertificate',
                     'CloudServicesMailEnabled', 'TreatMessagesAsInternal', 'EFSkipLastIP', 'EFSkipIPs', 'EFUsers', 'EFTestMode')
        $inbound = @(); $inFailed = $false
        try { $inbound = @(Get-SharedInboundConnectors) } catch { $inFailed = $true; Add-LogEntry -Section 'Inbound connectors' -ErrorRecord $_ }
        $inRows = @($inbound | ForEach-Object {
            $efOn = (Test-IsTrue $_.EFSkipLastIP) -or ((Get-NonEmptyCount $_.EFSkipIPs) -gt 0)
            $o = [ordered]@{ Name = $_.Name; EnhancedFiltering = if ($efOn) { 'On' } else { 'Off' } }
            foreach ($p in $inProps) { $o[$p] = $_.$p }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'InboundConnectors' -Description 'Inbound connectors' -Rows $inRows -Columns @('Name','ConnectorType','Enabled','EnhancedFiltering') -Failed:$inFailed
        foreach ($c in $inbound) {
            Add-Line "### $(ConvertTo-MdHeading $c.Name)"
            Add-Line
            $efOn = (Test-IsTrue $c.EFSkipLastIP) -or ((Get-NonEmptyCount $c.EFSkipIPs) -gt 0)
            Out-Card -Settings ([ordered]@{
                'Connector type'              = $c.ConnectorType
                'Enabled'                     = $c.Enabled
                'Sender domains'              = @{ Value = (Format-ConnectorDomain $c.SenderDomains); CsvFile = 'InboundConnectors.csv' }
                'Sender IP addresses'         = @{ Value = $c.SenderIPAddresses; CsvFile = 'InboundConnectors.csv' }
                'Require TLS'                 = $c.RequireTls
                'TLS sender certificate'      = $c.TlsSenderCertificateName
                'Restrict domains to IPs'     = $c.RestrictDomainsToIPAddresses
                'Restrict domains to cert'    = $c.RestrictDomainsToCertificate
                'Cloud services mail enabled' = $c.CloudServicesMailEnabled
                'Treat as internal'           = $c.TreatMessagesAsInternal
                'Enhanced Filtering'          = $(if ($efOn) { 'On' } else { 'Off' })
                'EF skip last IP'             = $c.EFSkipLastIP
                'EF skip IPs'                 = @{ Value = $c.EFSkipIPs; CsvFile = 'InboundConnectors.csv' }
                'EF users'                    = @{ Value = $c.EFUsers; CsvFile = 'InboundConnectors.csv' }
                'EF test mode'                = $c.EFTestMode
            })
        }

        Add-Line "## Outbound connectors"
        Add-Line
        $outProps = @('ConnectorType', 'Enabled', 'RecipientDomains', 'SmartHosts', 'TlsSettings', 'TlsDomain',
                      'UseMXRecord', 'CloudServicesMailEnabled', 'IsTransportRuleScoped', 'RouteAllMessagesViaOnPremises')
        $outbound = @(); $outFailed = $false
        try { $outbound = @(Get-SharedOutboundConnectors) } catch { $outFailed = $true; Add-LogEntry -Section 'Outbound connectors' -ErrorRecord $_ }
        $outRows = @($outbound | ForEach-Object {
            $o = [ordered]@{ Name = $_.Name }
            foreach ($p in $outProps) { $o[$p] = $_.$p }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'OutboundConnectors' -Description 'Outbound connectors' -Rows $outRows -Columns @('Name','ConnectorType','Enabled','UseMXRecord') -Failed:$outFailed
        foreach ($c in $outbound) {
            Add-Line "### $(ConvertTo-MdHeading $c.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Connector type'              = $c.ConnectorType
                'Enabled'                     = $c.Enabled
                'Recipient domains'           = @{ Value = (Format-ConnectorDomain $c.RecipientDomains); CsvFile = 'OutboundConnectors.csv' }
                'Smart hosts'                 = @{ Value = $c.SmartHosts; CsvFile = 'OutboundConnectors.csv' }
                'TLS settings'                = $c.TlsSettings
                'TLS domain'                  = $c.TlsDomain
                'Use MX record'               = $c.UseMXRecord
                'Cloud services mail enabled' = $c.CloudServicesMailEnabled
                'Scoped by transport rule'    = $c.IsTransportRuleScoped
                'Route all via on-premises'   = $c.RouteAllMessagesViaOnPremises
            })
        }

        Add-Line "## Transport rules"
        Add-Line
        $ruleProps = @('State', 'Priority', 'Mode', 'Description', 'Comments', 'SenderAddressLocation', 'WhenChanged')
        $rules = @(); $trFailed = $false
        try { $rules = @(Get-SharedTransportRules) } catch { $trFailed = $true; Add-LogEntry -Section 'Transport rules' -ErrorRecord $_ }
        $ruleRows = @($rules | ForEach-Object {
            $o = [ordered]@{ Name = $_.Name }
            foreach ($p in $ruleProps) { $o[$p] = $_.$p }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'TransportRules' -Description 'Transport rules' -Rows $ruleRows -Columns @('Priority','Name','State','Mode') -Failed:$trFailed
        foreach ($r in $rules) {
            Add-Line "### $(ConvertTo-MdHeading $r.Name)"
            Add-Line
            $desc = "$($r.Description)"; if ($desc.Length -gt 500) { $desc = $desc.Substring(0, 500) + ' … (full text in TransportRules.csv)' }
            $cmts = "$($r.Comments)"; if ($cmts.Length -gt 500) { $cmts = $cmts.Substring(0, 500) + ' … (full text in TransportRules.csv)' }
            Out-Card -Settings ([ordered]@{
                'State'                    = $r.State
                'Priority'                 = $r.Priority
                'Mode'                     = $r.Mode
                'Description'              = $desc
                'Comments'                 = $cmts
                'Sender address location'  = $r.SenderAddressLocation
                'Last changed'             = $r.WhenChanged
            })
        }

        Add-Line "## Remote domains"
        Add-Line
        $remote = @(); $remoteFailed = $false
        try { $remote = @(Get-RemoteDomain) } catch { $remoteFailed = $true; Add-LogEntry -Section 'Remote domains' -ErrorRecord $_ }
        $remoteRows = @($remote | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name; DomainName = $_.DomainName; AutoForwardEnabled = $_.AutoForwardEnabled; AutoReplyEnabled = $_.AutoReplyEnabled
                AllowedOOFType = $_.AllowedOOFType; TNEFEnabled = $_.TNEFEnabled; CharacterSet = $_.CharacterSet
                DeliveryReportEnabled = $_.DeliveryReportEnabled; NDREnabled = $_.NDREnabled
                MeetingForwardNotificationEnabled = $_.MeetingForwardNotificationEnabled; NonMimeCharacterSet = $_.NonMimeCharacterSet
            }
        })
        Out-Overview -CsvName 'RemoteDomains' -Description 'Remote domains' -Rows $remoteRows -Columns @('Name','DomainName','AutoForwardEnabled','AllowedOOFType') -Failed:$remoteFailed
        foreach ($rd in $remote) {
            Add-Line "### $(ConvertTo-MdHeading "$($rd.Name) ($($rd.DomainName))")"
            Add-Line
            $tnef = if ($null -eq $rd.TNEFEnabled -or "$($rd.TNEFEnabled)" -eq '') { 'Follow user settings' } elseif (Test-IsTrue $rd.TNEFEnabled) { 'Always' } else { 'Never' }
            Out-Card -Settings ([ordered]@{
                'Automatic replies'              = $rd.AllowedOOFType
                'Automatic forwarding'           = $rd.AutoForwardEnabled
                'Automatic replies enabled'      = $rd.AutoReplyEnabled
                'Delivery reports'               = $rd.DeliveryReportEnabled
                'Non-delivery reports'           = $rd.NDREnabled
                'Meeting forward notifications'  = $rd.MeetingForwardNotificationEnabled
                'Use rich-text format (TNEF)'    = $tnef
                'MIME character set'             = $rd.CharacterSet
                'Non-MIME character set'         = $rd.NonMimeCharacterSet
            })
        }

        Add-Line "## Journal rules"
        Add-Line
        $journal = @(); $journalFailed = $false
        try { $journal = @(Get-JournalRule) } catch { $journalFailed = $true; Add-LogEntry -Section 'Journal rules' -ErrorRecord $_ }
        $journalRows = @($journal | ForEach-Object {
            $scopeText = switch ("$($_.Scope)") {
                'Global'   { 'All messages' }
                'Internal' { 'Internal messages only' }
                'External' { 'External messages only' }
                default    { "$($_.Scope)" }
            }
            [PSCustomObject]@{ Name = $_.Name; JournalEmailAddress = $_.JournalEmailAddress; Scope = $scopeText; Recipient = $(if ("$($_.Recipient)" -match '\S') { $_.Recipient } else { 'All messages' }); Enabled = $_.Enabled }
        })
        Out-Overview -CsvName 'JournalRules' -Description 'Journal rules' -Rows $journalRows -Columns @('Name','JournalEmailAddress','Scope','Enabled') -Failed:$journalFailed

        Add-Line "## High Volume Email (HVE) accounts"
        Add-Line
        $hve = @(); $hveFailed = $false
        try { $hve = @(Get-MailUser -HVEAccount -ResultSize Unlimited) } catch { $hveFailed = $true; Add-LogEntry -Section 'HVE accounts' -ErrorRecord $_ }
        $hveRows = @($hve | ForEach-Object {
            [PSCustomObject]@{ DisplayName = $_.DisplayName; PrimarySmtpAddress = $_.PrimarySmtpAddress; MaxSendPerMinute = $_.MaxSendPerMinute }
        })
        Out-Overview -CsvName 'HveAccounts' -Description 'HVE accounts' -Rows $hveRows -Columns @('DisplayName','PrimarySmtpAddress','MaxSendPerMinute') -Failed:$hveFailed
    }
}

function Write-RecipientsSection {
    Invoke-Section -Title 'Recipients' -Body {
        Add-Line "Mailbox and recipient inventory: counts by type, shared/room/equipment mailboxes, inactive and soft-deleted mailboxes, and a mailbox feature summary."
        Add-Line

        $mbx = @(); $mbxFailed = $false
        try { $mbx = @(Get-SharedMailboxes) } catch { $mbxFailed = $true; Add-LogEntry -Section 'Mailboxes' -ErrorRecord $_ }

        Add-Line "## Counts by type"
        Add-Line
        $mailUsers = @(); $contacts = @(); $muFailed = $false; $mcFailed = $false
        try { $mailUsers = @(Get-MailUser -ResultSize Unlimited) } catch { $muFailed = $true; Add-LogEntry -Section 'Mail users' -ErrorRecord $_ }
        try { $contacts = @(Get-MailContact -ResultSize Unlimited) } catch { $mcFailed = $true; Add-LogEntry -Section 'Mail contacts' -ErrorRecord $_ }
        $typeRows = @($mbx | Group-Object RecipientTypeDetails | Sort-Object Name | ForEach-Object {
            [PSCustomObject]@{ Type = $_.Name; Count = $_.Count }
        })
        if ($mbxFailed) { $typeRows += [PSCustomObject]@{ Type = 'Mailbox (all types)'; Count = $script:LookupFailedText } }
        $typeRows += [PSCustomObject]@{ Type = 'MailUser'; Count = $(if ($muFailed) { $script:LookupFailedText } else { @($mailUsers | Where-Object { $_.RecipientTypeDetails -eq 'MailUser' }).Count }) }
        $typeRows += [PSCustomObject]@{ Type = 'GuestMailUser'; Count = $(if ($muFailed) { $script:LookupFailedText } else { @($mailUsers | Where-Object { $_.RecipientTypeDetails -eq 'GuestMailUser' }).Count }) }
        $typeRows += [PSCustomObject]@{ Type = 'MailContact'; Count = $(if ($mcFailed) { $script:LookupFailedText } else { $contacts.Count }) }
        Out-Overview -CsvName 'RecipientCounts' -Description 'Recipient counts by type' -Rows $typeRows -Columns @('Type','Count')
        Register-Csv -Name 'Mailboxes' -Description 'All mailboxes with collected properties' -Rows $mbx -Failed:$mbxFailed
        Register-Csv -Name 'MailUsers' -Description 'Mail users' -Rows $mailUsers -Failed:$muFailed
        Register-Csv -Name 'MailContacts' -Description 'Mail contacts' -Rows $contacts -Failed:$mcFailed

        Add-Line "## Shared, room and equipment mailboxes"
        Add-Line
        $shared = @($mbx | Where-Object { $_.RecipientTypeDetails -in @('SharedMailbox','RoomMailbox','EquipmentMailbox') } | ForEach-Object {
            [PSCustomObject]@{ Name = $_.DisplayName; Email = $_.PrimarySmtpAddress; Type = $_.RecipientTypeDetails }
        })
        Out-Overview -CsvName 'SharedRoomEquipmentMailboxes' -Description 'Shared, room and equipment mailboxes' -Rows $shared -Columns @('Name','Email','Type') -Failed:$mbxFailed

        Add-Line "## Inactive and soft-deleted mailboxes"
        Add-Line
        $inactive = @(); $softDeleted = @(); $inactiveFailed = $false; $softFailed = $false
        try { $inactive = @(Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited) } catch { $inactiveFailed = $true; Add-LogEntry -Section 'Inactive mailboxes' -ErrorRecord $_ }
        try { $softDeleted = @(Get-EXOMailbox -SoftDeletedMailbox -ResultSize Unlimited) } catch { $softFailed = $true; Add-LogEntry -Section 'Soft-deleted mailboxes' -ErrorRecord $_ }
        Register-Csv -Name 'InactiveMailboxes' -Description 'Inactive mailboxes' -Rows $inactive -Failed:$inactiveFailed
        Register-Csv -Name 'SoftDeletedMailboxes' -Description 'Soft-deleted mailboxes' -Rows $softDeleted -Failed:$softFailed
        Out-Card -Settings ([ordered]@{
            'Inactive mailboxes'     = $(if ($inactiveFailed) { $script:LookupFailedText } else { $inactive.Count })
            'Soft-deleted mailboxes' = $(if ($softFailed) { $script:LookupFailedText } else { $softDeleted.Count })
        })

        Add-Line "## Mailbox feature summary"
        Add-Line
        Out-Card -Settings ([ordered]@{
            'Litigation hold enabled'      = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { Test-IsTrue $_.LitigationHoldEnabled }).Count })
            'Archive enabled'              = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count })
            'Auto-expanding archive'       = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { Test-IsTrue $_.AutoExpandingArchiveEnabled }).Count })
            'Retention hold enabled'       = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { Test-IsTrue $_.RetentionHoldEnabled }).Count })
            'Forwarding configured'        = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress }).Count })
        })

        Add-Line "## Mailboxes with forwarding"
        Add-Line
        $acceptedNames = @($script:AcceptedDomains | ForEach-Object { "$($_.DomainName)".ToLowerInvariant() })
        $fwdRows = @($mbx | Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress } | ForEach-Object {
            $fwdSmtp = "$($_.ForwardingSmtpAddress)" -replace '(?i)^smtp:', ''
            $external = 'No'
            if ($fwdSmtp) {
                if (-not $script:AcceptedDomains) { $external = 'Unknown' }
                else { $external = $(if ($acceptedNames -contains (($fwdSmtp -split '@')[-1]).ToLowerInvariant()) { 'No' } else { 'Yes' }) }
            }
            [PSCustomObject]@{ Name = $_.DisplayName; ForwardTo = $(if ($fwdSmtp) { $fwdSmtp } else { $_.ForwardingAddress }); External = $external; KeepCopy = $_.DeliverToMailboxAndForward }
        })
        Out-Overview -CsvName 'MailboxForwarding' -Description 'Mailboxes with forwarding configured' -Rows $fwdRows -Columns @('Name','ForwardTo','External','KeepCopy') -Failed:$mbxFailed

        Add-Line "## Mailbox plans (defaults for new mailboxes)"
        Add-Line
        $plans = @(); $plansFailed = $false
        try { $plans = @(Get-MailboxPlan) } catch { $plansFailed = $true; Add-LogEntry -Section 'MailboxPlan' -ErrorRecord $_ }
        $planRows = @($plans | ForEach-Object {
            [PSCustomObject]@{ DisplayName = $_.DisplayName; IsDefault = $_.IsDefault; ProhibitSendReceiveQuota = $_.ProhibitSendReceiveQuota
                               ProhibitSendQuota = $_.ProhibitSendQuota; IssueWarningQuota = $_.IssueWarningQuota
                               MaxSendSize = $_.MaxSendSize; MaxReceiveSize = $_.MaxReceiveSize; RetainDeletedItemsFor = $_.RetainDeletedItemsFor
                               RetentionPolicy = $_.RetentionPolicy; RoleAssignmentPolicy = $_.RoleAssignmentPolicy }
        })
        Out-Overview -CsvName 'MailboxPlans' -Description 'Mailbox plans' -Rows $planRows -Columns @('DisplayName','IsDefault','ProhibitSendReceiveQuota') -Failed:$plansFailed
        foreach ($pl in $planRows) {
            Add-Line "### $(ConvertTo-MdHeading $pl.DisplayName)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Default policy'             = $pl.IsDefault
                'Prohibit send and receive at' = $pl.ProhibitSendReceiveQuota
                'Prohibit send at'           = $pl.ProhibitSendQuota
                'Issue warning at'           = $pl.IssueWarningQuota
                'Max send size'              = $pl.MaxSendSize
                'Max receive size'           = $pl.MaxReceiveSize
                'Keep deleted items for'     = $pl.RetainDeletedItemsFor
                'Retention policy'           = $pl.RetentionPolicy
                'Role assignment policy'     = $pl.RoleAssignmentPolicy
            })
        }

        Add-Line "## Address lists and address book policies"
        Add-Line
        Add-Line "### Address book policies"
        Add-Line
        $abp = @(); $abpFailed = $false
        if (Test-CmdletAvailable Get-AddressBookPolicy) {
            try { $abp = @(Get-AddressBookPolicy) } catch { $abpFailed = $true; Add-LogEntry -Section 'AddressBookPolicy' -ErrorRecord $_ }
        }
        else {
            $reason = 'Get-AddressBookPolicy is not available to this account (requires the Address Lists role).'
            Add-Line "_Not available: $reason`_" ; Add-Line
            Add-LogEntry -Section 'AddressBookPolicy' -Reason $reason
        }
        $abpRows = @($abp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; GlobalAddressList = $_.GlobalAddressList; OfflineAddressBook = $_.OfflineAddressBook
                               RoomList = $_.RoomList; AddressLists = $_.AddressLists }
        })
        if (Test-CmdletAvailable Get-AddressBookPolicy) { Out-Overview -CsvName 'AddressBookPolicies' -Description 'Address book policies' -Rows $abpRows -Columns @('Name','GlobalAddressList','OfflineAddressBook','RoomList') -Failed:$abpFailed }
        else { Register-Csv -Name 'AddressBookPolicies' -Description 'Address book policies' -Rows $abpRows }
        Add-Line "### Address lists"
        Add-Line
        $al = @(); $alFailed = $false
        if (Test-CmdletAvailable Get-AddressList) {
            try { $al = @(Get-AddressList) } catch { $alFailed = $true; Add-LogEntry -Section 'AddressList' -ErrorRecord $_ }
        }
        else {
            $reason = 'Get-AddressList is not available to this account (requires the Address Lists role).'
            Add-Line "_Not available: $reason`_" ; Add-Line
            Add-LogEntry -Section 'AddressList' -Reason $reason
        }
        if (Test-CmdletAvailable Get-AddressList) { Out-Overview -CsvName 'AddressLists' -Description 'Address lists' -Rows @($al | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientFilter = $_.RecipientFilter }
        }) -Columns @('Name','RecipientFilter') -Failed:$alFailed }
        else { Register-Csv -Name 'AddressLists' -Description 'Address lists' -Rows $al }
        Add-Line "### Global address lists"
        Add-Line
        $gal = @(); $galFailed = $false
        if (Test-CmdletAvailable Get-GlobalAddressList) {
            try { $gal = @(Get-GlobalAddressList) } catch { $galFailed = $true; Add-LogEntry -Section 'GlobalAddressList' -ErrorRecord $_ }
        }
        else {
            $reason = 'Get-GlobalAddressList is not available to this account (requires the Address Lists role).'
            Add-Line "_Not available: $reason`_" ; Add-Line
            Add-LogEntry -Section 'GlobalAddressList' -Reason $reason
        }
        if (Test-CmdletAvailable Get-GlobalAddressList) { Out-Overview -CsvName 'GlobalAddressLists' -Description 'Global address lists' -Rows @($gal | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; RecipientFilter = $_.RecipientFilter }
        }) -Columns @('Name','RecipientFilter') -Failed:$galFailed }
        else { Register-Csv -Name 'GlobalAddressLists' -Description 'Global address lists' -Rows $gal }
        Add-Line "### Offline address books"
        Add-Line
        $oab = @(); $oabFailed = $false
        if (Test-CmdletAvailable Get-OfflineAddressBook) {
            try { $oab = @(Get-OfflineAddressBook) } catch { $oabFailed = $true; Add-LogEntry -Section 'OfflineAddressBook' -ErrorRecord $_ }
        }
        else {
            $reason = 'Get-OfflineAddressBook is not available to this account (requires the Address Lists role).'
            Add-Line "_Not available: $reason`_" ; Add-Line
            Add-LogEntry -Section 'OfflineAddressBook' -Reason $reason
        }
        if (Test-CmdletAvailable Get-OfflineAddressBook) { Out-Overview -CsvName 'OfflineAddressBooks' -Description 'Offline address books' -Rows @($oab | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; IsDefault = $_.IsDefault; AddressLists = $_.AddressLists }
        }) -Columns @('Name','IsDefault','AddressLists') -Failed:$oabFailed }
        else { Register-Csv -Name 'OfflineAddressBooks' -Description 'Offline address books' -Rows $oab }

        Add-Line "## Public folders"
        Add-Line
        $pfMbx = @(); $pfFailed = $false
        try { $pfMbx = @(Get-Mailbox -PublicFolder -ResultSize Unlimited) } catch { $pfFailed = $true; Add-LogEntry -Section 'Public folder mailboxes' -ErrorRecord $_ }
        $pfOrg = $script:OrgConfig
        if (-not $pfOrg) { try { $pfOrg = Get-OrganizationConfig } catch { } }
        Out-Card -Settings ([ordered]@{
            'Public folders enabled'      = $pfOrg.PublicFoldersEnabled
            'Root public folder mailbox'  = $pfOrg.RootPublicFolderMailbox
            'Public folder mailboxes'     = $(if ($pfFailed) { $script:LookupFailedText } else { $pfMbx.Count })
        })

        Add-Line "## Room calendar processing"
        Add-Line
        $rooms = @($mbx | Where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox' })
        if ($rooms.Count -gt 200) { Add-LogEntry -Section 'Room calendar processing' -Reason "Only the first 200 of $($rooms.Count) room mailboxes were processed"; $rooms = @($rooms | Select-Object -First 200) }
        $calRows = @()
        foreach ($room in $rooms) {
            try {
                $cp = Get-CalendarProcessing -Identity "$($room.PrimarySmtpAddress)"
                $calRows += [PSCustomObject]@{ Room = $room.DisplayName; AutomateProcessing = $cp.AutomateProcessing
                    BookingWindowInDays = $cp.BookingWindowInDays; MaximumDurationInMinutes = $cp.MaximumDurationInMinutes
                    MaxDuration = $(if ("$($cp.MaximumDurationInMinutes)" -eq '0') { 'No limit' } else { $cp.MaximumDurationInMinutes })
                    AllowConflicts = $cp.AllowConflicts; AllowRecurringMeetings = $cp.AllowRecurringMeetings
                    ProcessExternalMeetingMessages = $cp.ProcessExternalMeetingMessages; AddOrganizerToSubject = $cp.AddOrganizerToSubject
                    DeleteSubject = $cp.DeleteSubject; AllBookInPolicy = $cp.AllBookInPolicy }
            }
            catch { Add-LogEntry -Section "CalendarProcessing ($($room.PrimarySmtpAddress))" -ErrorRecord $_ }
        }
        Out-Overview -CsvName 'RoomCalendarProcessing' -Description 'Calendar processing for room mailboxes' -Rows $calRows -Columns @('Room','AutomateProcessing','BookingWindowInDays','MaxDuration')
    }
}

function Write-GroupsSection {
    Invoke-Section -Title 'Groups' -Body {
        Add-Line "Distribution groups, mail-enabled security groups, dynamic distribution groups, Microsoft 365 groups and room lists."
        Add-Line

        $dg = @(); $dyn = @(); $m365 = @(); $dgFailed = $false; $dynFailed = $false; $m365Failed = $false
        try { $dg = @(Get-SharedDistributionGroups) } catch { $dgFailed = $true; Add-LogEntry -Section 'Distribution groups' -ErrorRecord $_ }
        try { $dyn = @(Get-SharedDynamicGroups) } catch { $dynFailed = $true; Add-LogEntry -Section 'Dynamic distribution groups' -ErrorRecord $_ }
        try { $m365 = @(Get-SharedM365Groups) } catch { $m365Failed = $true; Add-LogEntry -Section 'Microsoft 365 groups' -ErrorRecord $_ }

        Add-Line "## Counts"
        Add-Line
        Out-Card -Settings ([ordered]@{
            'Distribution groups'         = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup' }).Count })
            'Mail-enabled security groups' = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup' }).Count })
            'Dynamic distribution groups' = $(if ($dynFailed) { $script:LookupFailedText } else { $dyn.Count })
            'Microsoft 365 groups'        = $(if ($m365Failed) { $script:LookupFailedText } else { $m365.Count })
            'Room lists'                  = $(if ($dgFailed) { $script:LookupFailedText } else { @($dg | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' }).Count })
        })
        Register-Csv -Name 'DistributionGroups' -Description 'Distribution and mail-enabled security groups' -Rows $dg -Failed:$dgFailed
        Register-Csv -Name 'M365Groups' -Description 'Microsoft 365 groups' -Rows $m365 -Failed:$m365Failed

        Add-Line "## Microsoft 365 groups"
        Add-Line
        Out-Card -Settings ([ordered]@{
            'Private'                              = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { $_.AccessType -eq 'Private' }).Count })
            'Public'                               = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { $_.AccessType -eq 'Public' }).Count })
            'Teams-enabled'                        = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { @($_.ResourceProvisioningOptions) -contains 'Team' }).Count })
            'Hidden from the address list'         = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { Test-IsTrue $_.HiddenFromAddressListsEnabled }).Count })
            'Accepting external senders'           = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { -not (Test-IsTrue $_.RequireSenderAuthenticationEnabled) }).Count })
            'With guests'                          = $(if ($m365Failed) { $script:LookupFailedText } else { @($m365 | Where-Object { $_.GroupExternalMemberCount -gt 0 }).Count })
        })

        Add-Line "## Groups without owners"
        Add-Line
        $noOwner = @($dg | Where-Object { $_.RecipientTypeDetails -ne 'RoomList' -and (-not $_.ManagedBy -or @($_.ManagedBy).Count -eq 0) } | ForEach-Object {
            [PSCustomObject]@{ Name = $(if ("$($_.DisplayName)" -match '\S') { $_.DisplayName } else { $_.Name }); Type = $_.RecipientTypeDetails; Email = $_.PrimarySmtpAddress }
        })
        $noOwner += @($m365 | Where-Object { -not $_.ManagedBy -or @($_.ManagedBy).Count -eq 0 } | ForEach-Object {
            [PSCustomObject]@{ Name = $(if ("$($_.DisplayName)" -match '\S') { $_.DisplayName } else { $_.Name }); Type = 'GroupMailbox'; Email = $_.PrimarySmtpAddress }
        })
        Out-Overview -CsvName 'GroupsNoOwners' -Description 'Groups without owners' -Rows $noOwner -Columns @('Name','Type','Email') -Failed:($dgFailed -and $m365Failed)
        if ($dgFailed -xor $m365Failed) { Add-Line "_Incomplete: $(if ($dgFailed) { 'distribution groups' } else { 'Microsoft 365 groups' }) could not be read (see collection log)._"; Add-Line }

        Add-Line "## Groups accepting mail from external senders"
        Add-Line
        $extSenders = @($dg | Where-Object { $_.RecipientTypeDetails -ne 'RoomList' -and -not (Test-IsTrue $_.RequireSenderAuthenticationEnabled) } | ForEach-Object {
            [PSCustomObject]@{ Name = $(if ("$($_.DisplayName)" -match '\S') { $_.DisplayName } else { $_.Name }); Type = $_.RecipientTypeDetails; Email = $_.PrimarySmtpAddress }
        })
        $extSenders += @($m365 | Where-Object { -not (Test-IsTrue $_.RequireSenderAuthenticationEnabled) } | ForEach-Object {
            [PSCustomObject]@{ Name = $(if ("$($_.DisplayName)" -match '\S') { $_.DisplayName } else { $_.Name }); Type = 'GroupMailbox'; Email = $_.PrimarySmtpAddress }
        })
        Out-Overview -CsvName 'GroupsAcceptExternal' -Description 'Groups accepting mail from external senders' -Rows $extSenders -Columns @('Name','Type','Email') -Failed:($dgFailed -and $m365Failed)
        if ($dgFailed -xor $m365Failed) { Add-Line "_Incomplete: $(if ($dgFailed) { 'distribution groups' } else { 'Microsoft 365 groups' }) could not be read (see collection log)._"; Add-Line }

        Add-Line "## Dynamic distribution groups"
        Add-Line
        Register-Csv -Name 'DynamicDistributionGroups' -Description 'Dynamic distribution groups' -Rows $dyn -Failed:$dynFailed
        if ($dynFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($dyn.Count -eq 0) { Add-Line "_None found._"; Add-Line }
        foreach ($g in $dyn) {
            Add-Line "### $(ConvertTo-MdHeading $g.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Primary SMTP address' = $g.PrimarySmtpAddress
                'Recipient filter'     = $g.RecipientFilter
                'Managed by'           = @{ Value = $g.ManagedBy; CsvFile = 'DynamicDistributionGroups.csv' }
            })
        }

        Add-Line "## Room lists"
        Add-Line
        $roomLists = @($dg | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' } | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Email = $_.PrimarySmtpAddress }
        })
        Out-Overview -CsvName 'RoomLists' -Description 'Room list distribution groups' -Rows $roomLists -Columns @('Name','Email') -Failed:$dgFailed
    }
}

function Write-ClientAccessSection {
    Invoke-Section -Title 'Client access and mailbox policies' -Body {
        Add-Line "How clients connect to mailboxes: default role assignment, OWA, mobile device and authentication policies, plus protocol enablement across mailboxes."
        Add-Line

        Add-Line "## Mailbox policy assignment"
        Add-Line
        $mbx = @(); $mbxFailed = $false
        try { $mbx = @(Get-SharedMailboxes) } catch { $mbxFailed = $true; Add-LogEntry -Section 'Mailboxes' -ErrorRecord $_ }
        $cas = @(); $casFailed = $false
        try { $cas = @(Get-SharedCasMailboxes) } catch { $casFailed = $true; Add-LogEntry -Section 'CAS mailboxes' -ErrorRecord $_ }
        Add-Line "### Role assignment policy"
        Add-Line
        $rapSpread = @($mbx | Group-Object RoleAssignmentPolicy | Sort-Object Count -Descending | ForEach-Object {
            [PSCustomObject]@{ Policy = $(if ("$($_.Name)" -match '\S') { $_.Name } else { '(default policy)' }); Mailboxes = $_.Count }
        })
        Out-Overview -CsvName 'MailboxPolicySpread_RoleAssignment' -Description 'Role assignment policy usage across mailboxes' -Rows $rapSpread -Columns @('Policy','Mailboxes') -Failed:$mbxFailed
        Add-Line "### Retention policy"
        Add-Line
        $retSpread = @($mbx | Group-Object RetentionPolicy | Sort-Object Count -Descending | ForEach-Object {
            [PSCustomObject]@{ Policy = $(if ("$($_.Name)" -match '\S') { $_.Name } else { '(no retention policy)' }); Mailboxes = $_.Count }
        })
        Out-Overview -CsvName 'MailboxPolicySpread_Retention' -Description 'Retention policy usage across mailboxes' -Rows $retSpread -Columns @('Policy','Mailboxes') -Failed:$mbxFailed
        Add-Line "### OWA mailbox policy"
        Add-Line
        $owaSpread = @($cas | Group-Object OwaMailboxPolicy | Sort-Object Count -Descending | ForEach-Object {
            [PSCustomObject]@{ Policy = $(if ("$($_.Name)" -match '\S') { $_.Name } else { '(default policy)' }); Mailboxes = $_.Count }
        })
        Out-Overview -CsvName 'MailboxPolicySpread_Owa' -Description 'OWA mailbox policy usage across mailboxes' -Rows $owaSpread -Columns @('Policy','Mailboxes') -Failed:$casFailed
        Add-Line "### Mobile device mailbox policy"
        Add-Line
        $mdpSpread = @($cas | Group-Object ActiveSyncMailboxPolicy | Sort-Object Count -Descending | ForEach-Object {
            [PSCustomObject]@{ Policy = $(if ("$($_.Name)" -match '\S') { $_.Name } else { '(default policy)' }); Mailboxes = $_.Count }
        })
        Out-Overview -CsvName 'MailboxPolicySpread_MobileDevice' -Description 'Mobile device mailbox policy usage across mailboxes' -Rows $mdpSpread -Columns @('Policy','Mailboxes') -Failed:$casFailed

        Add-Line "## Protocol enablement counts"
        Add-Line
        $orgSmtpAuth = $null
        try { $orgSmtpAuth = (@(Get-SharedTransportConfig) | Select-Object -First 1).SmtpClientAuthenticationDisabled } catch { }
        $smtpOrgText = if ($null -eq $orgSmtpAuth) { 'unknown' } elseif (Test-IsTrue $orgSmtpAuth) { 'disabled' } else { 'enabled' }
        $smtpEnabled  = @($cas | Where-Object { "$($_.SmtpClientAuthenticationDisabled)" -eq 'False' }).Count
        $smtpDisabled = @($cas | Where-Object { "$($_.SmtpClientAuthenticationDisabled)" -eq 'True' }).Count
        $smtpFollows  = @($cas | Where-Object { $null -eq $_.SmtpClientAuthenticationDisabled -or "$($_.SmtpClientAuthenticationDisabled)" -eq '' }).Count
        $protoCount = {
            param([string]$Prop)
            if ($casFailed) { return $script:LookupFailedText }
            $en = @($cas | Where-Object { Test-IsTrue $_.$Prop }).Count
            $dis = @($cas | Where-Object { "$($_.$Prop)" -eq 'False' }).Count
            $un = @($cas | Where-Object { $null -eq $_.$Prop -or "$($_.$Prop)" -eq '' }).Count
            "$en / $dis" + $(if ($un -gt 0) { " / not set $un" } else { '' })
        }
        Register-Csv -Name 'CasMailboxes' -Description 'Client access settings per mailbox' -Rows $cas -Failed:$casFailed
        Out-Card -Settings ([ordered]@{
            'POP enabled / disabled'    = (& $protoCount 'PopEnabled')
            'IMAP enabled / disabled'   = (& $protoCount 'ImapEnabled')
            'EWS enabled / disabled'    = (& $protoCount 'EwsEnabled')
            'ActiveSync enabled / disabled' = (& $protoCount 'ActiveSyncEnabled')
            'MAPI enabled / disabled'   = (& $protoCount 'MapiEnabled')
            'OWA enabled / disabled'    = (& $protoCount 'OWAEnabled')
            'SMTP AUTH (per mailbox)'   = $(if ($casFailed) { $script:LookupFailedText } else { "enabled $smtpEnabled / disabled $smtpDisabled / follows org setting $smtpFollows (org-wide: SMTP AUTH $smtpOrgText)" })
        })

        Add-Line "## OWA mailbox policies"
        Add-Line
        $owaSettings = @(
            @{ Label = 'Default policy'; Prop = 'IsDefault' },
            @{ Label = 'Instant messaging'; Prop = 'InstantMessagingEnabled' },
            @{ Label = 'Email signature'; Prop = 'SignaturesEnabled' },
            @{ Label = 'Offline access'; Prop = 'AllowOfflineOn' },
            @{ Label = 'Personal accounts (new Outlook)'; Prop = 'PersonalAccountsEnabled' },
            @{ Label = 'Personal account calendars'; Prop = 'PersonalAccountCalendarsEnabled' },
            @{ Label = 'New Outlook for Windows allowed'; Prop = 'OneWinNativeOutlookEnabled' },
            @{ Label = 'Show the new Outlook toggle'; Prop = 'OutlookBetaToggleEnabled' },
            @{ Label = 'Classic attachments'; Prop = 'ClassicAttachmentsEnabled' },
            @{ Label = 'Third-party file providers'; Prop = 'ThirdPartyFileProvidersEnabled' },
            @{ Label = 'Additional storage providers'; Prop = 'AdditionalStorageProvidersAvailable' },
            @{ Label = 'Direct file access (public computers)'; Prop = 'DirectFileAccessOnPublicComputersEnabled' },
            @{ Label = 'Direct file access (private computers)'; Prop = 'DirectFileAccessOnPrivateComputersEnabled' },
            @{ Label = 'Web-ready viewing (public computers)'; Prop = 'WacViewingOnPublicComputersEnabled' },
            @{ Label = 'Web-ready viewing (private computers)'; Prop = 'WacViewingOnPrivateComputersEnabled' },
            @{ Label = 'External image proxy'; Prop = 'ExternalImageProxyEnabled' },
            @{ Label = 'Conditional access policy'; Prop = 'ConditionalAccessPolicy' }
        )
        $owa = @(); $owaFailed = $false
        try { $owa = @(Get-OwaMailboxPolicy) } catch { $owaFailed = $true; Add-LogEntry -Section 'OwaMailboxPolicy' -ErrorRecord $_ }
        $owaRows = @($owa | ForEach-Object {
            $o = [ordered]@{ Name = $_.Name }
            foreach ($s in $owaSettings) { $o[$s.Prop] = $_.($s.Prop) }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'OwaMailboxPolicies' -Description 'OWA mailbox policies' -Rows $owaRows -Columns @('Name','IsDefault','ConditionalAccessPolicy') -Failed:$owaFailed
        foreach ($pol in $owa) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            $card = [ordered]@{}
            foreach ($s in $owaSettings) { $card[$s.Label] = @{ Value = $pol.($s.Prop); CsvFile = 'OwaMailboxPolicies.csv' } }
            Out-Card -Settings $card
        }

        Add-Line "## Mobile device mailbox policies"
        Add-Line
        $mdpSettings = @(
            @{ Label = 'Default policy'; Prop = 'IsDefault' },
            @{ Label = "Allow mobile devices that don't fully support these policies to synchronize"; Prop = 'AllowNonProvisionableDevices' },
            @{ Label = 'Require a password'; Prop = 'PasswordEnabled' },
            @{ Label = 'Allow simple passwords'; Prop = 'AllowSimplePassword' },
            @{ Label = 'Require an alphanumeric password'; Prop = 'AlphanumericPasswordRequired' },
            @{ Label = 'Minimum password length'; Prop = 'MinPasswordLength' },
            @{ Label = 'Number of sign-in failures before the device is wiped'; Prop = 'MaxPasswordFailedAttempts' },
            @{ Label = 'Require sign-in after the device has been inactive for'; Prop = 'MaxInactivityTimeLock' },
            @{ Label = 'Enforce password lifetime'; Prop = 'PasswordExpiration' },
            @{ Label = 'Password recycle count'; Prop = 'PasswordHistory' },
            @{ Label = 'Require encryption on the device'; Prop = 'RequireDeviceEncryption' }
        )
        $mdp = @(); $mdpFailed = $false
        try { $mdp = @(Get-MobileDeviceMailboxPolicy) } catch { $mdpFailed = $true; Add-LogEntry -Section 'MobileDeviceMailboxPolicy' -ErrorRecord $_ }
        $mdpRows = @($mdp | ForEach-Object {
            $o = [ordered]@{ Name = $_.Name }
            foreach ($s in $mdpSettings) { $o[$s.Prop] = $_.($s.Prop) }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'MobileDeviceMailboxPolicies' -Description 'Mobile device (ActiveSync) mailbox policies' -Rows $mdpRows -Columns @('Name','IsDefault','PasswordEnabled') -Failed:$mdpFailed
        foreach ($pol in $mdp) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            $card = [ordered]@{}
            foreach ($s in $mdpSettings) { $card[$s.Label] = $pol.($s.Prop) }
            Out-Card -Settings $card
        }

        Add-Line "## ActiveSync organization settings and access rules"
        Add-Line
        $aso = @(); $asoFailed = $false
        try { $aso = @(Get-ActiveSyncOrganizationSettings) } catch { $asoFailed = $true; Add-LogEntry -Section 'ActiveSyncOrganizationSettings' -ErrorRecord $_ }
        if ($aso.Count -gt 0) {
            Out-Card -Settings ([ordered]@{
                'Default access level' = $aso[0].DefaultAccessLevel
                'User mail insert'     = $aso[0].UserMailInsert
                'Admin mail recipients' = @{ Value = $aso[0].AdminMailRecipients; CsvFile = 'ActiveSyncOrgSettings.csv' }
            })
        }
        elseif ($asoFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not configured._"; Add-Line }
        Register-Csv -Name 'ActiveSyncOrgSettings' -Description 'ActiveSync organization settings' -Rows $aso -Failed:$asoFailed
        $asr = @(); $asrFailed = $false
        try { $asr = @(Get-ActiveSyncDeviceAccessRule) } catch { $asrFailed = $true; Add-LogEntry -Section 'ActiveSyncDeviceAccessRule' -ErrorRecord $_ }
        $asrRows = @($asr | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Characteristic = $_.Characteristic; QueryString = $_.QueryString; AccessLevel = $_.AccessLevel }
        })
        Add-Line "### Device access rules"
        Add-Line
        Out-Overview -CsvName 'ActiveSyncDeviceAccessRules' -Description 'ActiveSync device access rules' -Rows $asrRows -Columns @('Name','Characteristic','QueryString','AccessLevel') -Failed:$asrFailed

        Add-Line "## Authentication policies"
        Add-Line
        $ap = @(); $apFailed = $false
        try { $ap = @(Get-AuthenticationPolicy) } catch { $apFailed = $true; Add-LogEntry -Section 'AuthenticationPolicy' -ErrorRecord $_ }
        $basicProps = [ordered]@{
            AllowBasicAuthPop = 'POP'; AllowBasicAuthImap = 'IMAP'; AllowBasicAuthSmtp = 'SMTP'
            AllowBasicAuthActiveSync = 'ActiveSync'; AllowBasicAuthWebService = 'Exchange Web Services'
            AllowBasicAuthMapi = 'MAPI'; AllowBasicAuthOfflineAddressBook = 'Offline Address Book'
            AllowBasicAuthRpc = 'RPC'; AllowBasicAuthPowershell = 'PowerShell'
            AllowBasicAuthAutodiscover = 'Autodiscover'; AllowBasicAuthOutlookService = 'Outlook service'
            AllowBasicAuthReportingWebServices = 'Reporting web services'
        }
        $apRows = @($ap | ForEach-Object {
            $pol = $_
            $o = [ordered]@{ Name = $pol.Name }
            foreach ($p in $basicProps.Keys) { $o[$p] = $pol.$p }
            foreach ($p in $pol.PSObject.Properties.Name | Where-Object { $_ -like 'BlockLegacyAuth*' }) { $o[$p] = $pol.$p }
            [PSCustomObject]$o
        })
        Register-Csv -Name 'AuthenticationPolicies' -Description 'Authentication policies' -Rows $apRows -Failed:$apFailed
        if ($apFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($ap.Count -eq 0) { Add-Line "_None found._"; Add-Line }
        foreach ($pol in $ap) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            $basicAllowed = @($basicProps.Keys | Where-Object { Test-IsTrue $pol.$_ } | ForEach-Object { $basicProps[$_] })
            $legacyBlocked = @($pol.PSObject.Properties.Name | Where-Object { $_ -like 'BlockLegacyAuth*' -and (Test-IsTrue $pol.$_) } | ForEach-Object { $_ -replace '^BlockLegacyAuth', '' })
            Out-Card -Settings ([ordered]@{
                'Basic authentication allowed for' = $(if ($basicAllowed.Count) { $basicAllowed -join ', ' } else { 'None' })
                'Legacy authentication blocked for' = $(if ($legacyBlocked.Count) { $legacyBlocked -join ', ' } else { 'None' })
            })
        }

        Add-Line "## Outlook add-ins in the Exchange organization catalog"
        Add-Line
        Add-Line "Organization-level Outlook add-ins stored in Exchange Online: Microsoft's default add-ins and add-ins installed through the retired classic Exchange admin center. These are no longer shown in any admin portal and are managed with Get-App / Set-App. Apps deployed through Microsoft 365 admin center > Settings > Integrated apps are listed in the next section."
        Add-Line
        $apps = @(); $appsFailed = $false
        try { $apps = @(Get-App -OrganizationApp) } catch { $appsFailed = $true; Add-LogEntry -Section 'OrganizationApps' -ErrorRecord $_ }
        $appRows = @($apps | ForEach-Object {
            [PSCustomObject]@{ DisplayName = $_.DisplayName; Enabled = $_.Enabled; DefaultStateForUser = $_.DefaultStateForUser; ProvidedTo = $_.ProvidedTo; AppId = $_.AppId; AppVersion = $_.AppVersion }
        })
        Out-Overview -CsvName 'OrganizationApps' -Description 'Outlook add-ins in the Exchange organization catalog (Get-App)' -Rows $appRows -Columns @('DisplayName','Enabled','DefaultStateForUser','ProvidedTo') -Failed:$appsFailed

        Add-Line "## Integrated apps (centrally deployed add-ins)"
        Add-Line
        if (-not $IncludeIntegratedApps) {
            Add-Line "_Skipped - run with -IncludeIntegratedApps to collect apps deployed through Microsoft 365 admin center > Settings > Integrated apps._"
            Add-Line
            Register-Csv -Name 'IntegratedApps' -Description 'Integrated apps / centrally deployed add-ins (Get-OrganizationAddIn)' -Rows @() -Note 'skipped'
        }
        elseif (-not $script:IntegratedAppsConnected) {
            Add-Line "_Not available: $($script:IntegratedAppsError)_"
            Add-Line
            Register-Csv -Name 'IntegratedApps' -Description 'Integrated apps / centrally deployed add-ins (Get-OrganizationAddIn)' -Rows @() -Failed
        }
        else {
            Add-Line "Collected with Get-OrganizationAddIn. Microsoft 365 apps that also run in Teams (for example Org Explorer, Planner, Bookings) may not be returned; compare with Microsoft 365 admin center > Settings > Integrated apps."
            Add-Line
            if ($script:IntegratedAppsListError) {
                Add-LogEntry -Section 'IntegratedApps' -ErrorRecord $script:IntegratedAppsListError
            }
            foreach ($k in $script:IntegratedAppsDetailErrors.Keys) {
                Add-LogEntry -Section "IntegratedApps detail ($k)" -ErrorRecord $script:IntegratedAppsDetailErrors[$k]
            }
            $iaRows = ConvertTo-IntegratedAppRows -List @($script:IntegratedAppsList) -Details $script:IntegratedAppsDetails
            Out-Overview -CsvName 'IntegratedApps' -Description 'Integrated apps / centrally deployed add-ins (Get-OrganizationAddIn)' -Rows $iaRows -Columns @('Name','Enabled','AssignedTo','ProductId') -Failed:([bool]$script:IntegratedAppsListError)
        }
    }
}

function Write-SharingSection {
    Invoke-Section -Title 'Sharing and federation' -Body {
        Add-Line "Calendar sharing and federation with other organisations: organization relationships, sharing policies and availability address spaces."
        Add-Line

        Add-Line "## Organization relationships"
        Add-Line
        $rel = @(); $relFailed = $false
        try { $rel = @(Get-OrganizationRelationship) } catch { $relFailed = $true; Add-LogEntry -Section 'OrganizationRelationship' -ErrorRecord $_ }
        $relRows = @($rel | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; DomainNames = $_.DomainNames; FreeBusyAccessEnabled = $_.FreeBusyAccessEnabled; FreeBusyAccessLevel = $_.FreeBusyAccessLevel; Enabled = $_.Enabled
                               MailTipsAccessEnabled = $_.MailTipsAccessEnabled; MailTipsAccessLevel = $_.MailTipsAccessLevel
                               DeliveryReportEnabled = $_.DeliveryReportEnabled; MailboxMoveEnabled = $_.MailboxMoveEnabled
                               MailboxMoveCapability = $_.MailboxMoveCapability; FreeBusyAccessScope = $_.FreeBusyAccessScope
                               TargetAutodiscoverEpr = $_.TargetAutodiscoverEpr; TargetApplicationUri = $_.TargetApplicationUri
                               OrganizationContact = $_.OrganizationContact }
        })
        Out-Overview -CsvName 'OrganizationRelationships' -Description 'Organization relationships' -Rows $relRows -Columns @('Name','Enabled','FreeBusyAccessEnabled','FreeBusyAccessLevel') -Failed:$relFailed
        foreach ($o in $rel) {
            Add-Line "### $(ConvertTo-MdHeading $o.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Enabled'                  = $o.Enabled
                'Domain names'             = @{ Value = $o.DomainNames; CsvFile = 'OrganizationRelationships.csv' }
                'Free/busy access enabled' = $o.FreeBusyAccessEnabled
                'Free/busy access level'   = $o.FreeBusyAccessLevel
                'Free/busy access scope'   = $o.FreeBusyAccessScope
                'MailTips access'          = $o.MailTipsAccessEnabled
                'MailTips access level'    = $o.MailTipsAccessLevel
                'Delivery reports'         = $o.DeliveryReportEnabled
                'Mailbox moves'            = $o.MailboxMoveEnabled
                'Mailbox move capability'  = $o.MailboxMoveCapability
                'Target Autodiscover endpoint' = $o.TargetAutodiscoverEpr
                'Target application URI'   = $o.TargetApplicationUri
                'Organization contact'     = $o.OrganizationContact
            })
        }

        Add-Line "## Sharing policies"
        Add-Line
        $sp = @(); $spFailed = $false
        try { $sp = @(Get-SharingPolicy) } catch { $spFailed = $true; Add-LogEntry -Section 'SharingPolicy' -ErrorRecord $_ }
        $spRows = @($sp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Domains = $_.Domains; SharingRules = (Format-SharingDomains $_.Domains); Default = $_.Default; Enabled = $_.Enabled }
        })
        Out-Overview -CsvName 'SharingPolicies' -Description 'Sharing policies' -Rows $spRows -Columns @('Name','Default','Enabled') -Failed:$spFailed
        foreach ($pol in $sp) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Default' = $pol.Default
                'Enabled' = $pol.Enabled
                'Sharing rules' = (Format-SharingDomains $pol.Domains)
            })
        }

        Add-Line "## Availability address spaces"
        Add-Line
        $aas = @(); $aasFailed = $false
        try { $aas = @(Get-AvailabilityAddressSpace) } catch { $aasFailed = $true; Add-LogEntry -Section 'AvailabilityAddressSpace' -ErrorRecord $_ }
        $aasRows = @($aas | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ForestName = $_.ForestName; AccessMethod = $_.AccessMethod }
        })
        Out-Overview -CsvName 'AvailabilityAddressSpaces' -Description 'Availability address spaces' -Rows $aasRows -Columns @('Name','ForestName','AccessMethod') -Failed:$aasFailed
        foreach ($a in $aas) {
            Add-Line "### $(ConvertTo-MdHeading $a.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Forest name'  = $a.ForestName
                'Access method' = $a.AccessMethod
            })
        }
    }
}

function Write-HybridSection {
    Invoke-Section -Title 'Hybrid and migration' -Body {
        Add-Line "Hybrid Exchange configuration and migration objects. If the tenant is cloud-only, these show as Not configured or empty."
        Add-Line

        Add-Line "## On-premises organization"
        Add-Line
        $opo = @(); $opoFailed = $false
        try { $opo = @(Get-OnPremisesOrganization) } catch { $opo = @(); $opoFailed = $true; Add-LogEntry -Section 'OnPremisesOrganization' -ErrorRecord $_ }
        if ($opoFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($opo.Count -eq 0) { Add-Line "_Not configured._"; Add-Line }
        foreach ($o in $opo) {
            Add-Line "### $(ConvertTo-MdHeading $o.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Organization name'        = $o.OrganizationName
                'Hybrid domains'           = @{ Value = $o.HybridDomains; CsvFile = 'OnPremisesOrganization.csv' }
                'Inbound connector'        = $o.InboundConnector
                'Outbound connector'       = $o.OutboundConnector
                'Organization relationship' = $o.OrganizationRelationship
            })
        }
        Register-Csv -Name 'OnPremisesOrganization' -Description 'On-premises organization (hybrid)' -Rows $opo -Failed:$opoFailed

        Add-Line "## Intra-organization connectors"
        Add-Line
        $ioc = @(); $iocFailed = $false
        try { $ioc = @(Get-IntraOrganizationConnector) } catch { $ioc = @(); $iocFailed = $true; Add-LogEntry -Section 'IntraOrganizationConnector' -ErrorRecord $_ }
        if ($iocFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($ioc.Count -eq 0) { Add-Line "_Not configured._"; Add-Line }
        foreach ($o in $ioc) {
            Add-Line "### $(ConvertTo-MdHeading $o.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Enabled'                = $o.Enabled
                'Target address domains' = @{ Value = $o.TargetAddressDomains; CsvFile = 'IntraOrganizationConnectors.csv' }
                'Discovery endpoint'     = $o.DiscoveryEndpoint
            })
        }
        Register-Csv -Name 'IntraOrganizationConnectors' -Description 'Intra-organization connectors' -Rows $ioc -Failed:$iocFailed

        Add-Line "## Migration endpoints and batches"
        Add-Line
        $mep = @(); $mepFailed = $false
        try { $mep = @(Get-MigrationEndpoint) } catch { $mep = @(); $mepFailed = $true; Add-LogEntry -Section 'MigrationEndpoint' -ErrorRecord $_ }
        if ($mepFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($mep.Count -eq 0) { Add-Line "_No migration endpoints._"; Add-Line }
        foreach ($o in $mep) {
            Add-Line "### $(ConvertTo-MdHeading $o.Identity)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Remote server'    = $o.RemoteServer
                'Exchange version' = $o.ExchangeVersion
            })
        }
        Register-Csv -Name 'MigrationEndpoints' -Description 'Migration endpoints' -Rows $mep -Failed:$mepFailed
        $batch = @(); $batchFailed = $false
        try { $batch = @(Get-MigrationBatch) } catch { $batch = @(); $batchFailed = $true; Add-LogEntry -Section 'MigrationBatch' -ErrorRecord $_ }
        $batchRows = @($batch | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Identity; Status = $_.Status; TotalCount = $_.TotalCount; FinalizedCount = $_.FinalizedCount }
        })
        Out-Overview -CsvName 'MigrationBatches' -Description 'Migration batches' -Rows $batchRows -Columns @('Name','Status','TotalCount','FinalizedCount') -Failed:$batchFailed
    }
}

function Write-EmailSecuritySection {
    Invoke-Section -Title 'Email security (EOP / Defender for Office 365)' -Body {
        Add-Line "Anti-phishing, anti-spam, anti-malware, Safe Attachments and Safe Links configuration, plus quarantine policies, advanced delivery and reporting settings."
        Add-Line

        Add-Line "## Preset security policies"
        Add-Line
        $presetRows = @()
        $eopPresetFailed = $false; $atpPresetFailed = $false
        $eop = @()
        try {
            $eop = @(Get-EOPProtectionPolicyRule)
            $presetRows += @($eop | ForEach-Object {
                [PSCustomObject]@{ Name = $_.Name; State = $_.State; Priority = $_.Priority; Type = 'EOP'; SentTo = $_.SentTo; SentToMemberOf = $_.SentToMemberOf; RecipientDomainIs = $_.RecipientDomainIs }
            })
        } catch { $eopPresetFailed = $true; Add-LogEntry -Section 'EOPProtectionPolicyRule' -ErrorRecord $_ }
        $atpPreset = @()
        try {
            $atpPreset = @(Get-ATPProtectionPolicyRule)
            $presetRows += @($atpPreset | ForEach-Object {
                [PSCustomObject]@{ Name = $_.Name; State = $_.State; Priority = $_.Priority; Type = 'ATP'; SentTo = $_.SentTo; SentToMemberOf = $_.SentToMemberOf; RecipientDomainIs = $_.RecipientDomainIs }
            })
        } catch { $atpPresetFailed = $true; Add-LogEntry -Section 'ATPProtectionPolicyRule' -ErrorRecord $_ }
        $presetRules = @($eop) + @($atpPreset)
        $presetPartFailed = $(if ($eopPresetFailed -xor $atpPresetFailed) { if ($eopPresetFailed) { 'EOP' } else { 'ATP (Defender for Office 365)' } })
        Out-Overview -CsvName 'PresetSecurityPolicies' -Description 'Preset security policies (EOP/ATP)' -Rows $presetRows -Columns @('Type','Name','State','Priority') -Failed:($eopPresetFailed -and $atpPresetFailed) -Note $(if ($presetPartFailed) { "incomplete: $presetPartFailed preset rules lookup failed" })
        if ($presetPartFailed) { Add-Line "_Incomplete: $presetPartFailed preset rules could not be read (see collection log)._"; Add-Line }
        $builtin = @(); $builtinFailed = $false
        try { $builtin = @(Get-ATPBuiltInProtectionRule) } catch { $builtinFailed = $true; Add-LogEntry -Section 'ATPBuiltInProtectionRule' -ErrorRecord $_ }
        $biRows = @($builtin | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; State = $_.State; SentTo = $_.SentTo; SentToMemberOf = $_.SentToMemberOf; RecipientDomainIs = $_.RecipientDomainIs }
        })
        Add-Line "### Built-in protection"
        Add-Line
        Out-Overview -CsvName 'BuiltInProtectionRule' -Description 'Built-in protection rule' -Rows $biRows -Columns @('Name','State') -Failed:$builtinFailed

        Add-Line "## Anti-phishing policies"
        Add-Line
        $phishSettings = @(
            @{ Label = 'Phishing email threshold'; Prop = 'PhishThresholdLevel' },
            @{ Label = 'Enable users to protect'; Prop = 'EnableTargetedUserProtection' },
            @{ Label = 'Users to protect'; Prop = 'TargetedUsersToProtect' },
            @{ Label = 'Include domains I own'; Prop = 'EnableOrganizationDomainsProtection' },
            @{ Label = 'Include custom domains'; Prop = 'EnableTargetedDomainsProtection' },
            @{ Label = 'Custom domains to protect'; Prop = 'TargetedDomainsToProtect' },
            @{ Label = 'Trusted senders'; Prop = 'ExcludedSenders' },
            @{ Label = 'Trusted domains'; Prop = 'ExcludedDomains' },
            @{ Label = 'Enable mailbox intelligence'; Prop = 'EnableMailboxIntelligence' },
            @{ Label = 'Enable intelligence for impersonation protection'; Prop = 'EnableMailboxIntelligenceProtection' },
            @{ Label = 'Enable spoof intelligence'; Prop = 'EnableSpoofIntelligence' },
            @{ Label = 'If a message is detected as user impersonation'; Prop = 'TargetedUserProtectionAction' },
            @{ Label = 'Quarantine policy (user impersonation)'; Prop = 'TargetedUserQuarantineTag' },
            @{ Label = 'If a message is detected as domain impersonation'; Prop = 'TargetedDomainProtectionAction' },
            @{ Label = 'Quarantine policy (domain impersonation)'; Prop = 'TargetedDomainQuarantineTag' },
            @{ Label = 'If mailbox intelligence detects an impersonated user'; Prop = 'MailboxIntelligenceProtectionAction' },
            @{ Label = 'Quarantine policy (mailbox intelligence)'; Prop = 'MailboxIntelligenceQuarantineTag' },
            @{ Label = 'Honor DMARC record policy'; Prop = 'HonorDmarcPolicy' },
            @{ Label = 'DMARC p=quarantine action'; Prop = 'DmarcQuarantineAction' },
            @{ Label = 'DMARC p=reject action'; Prop = 'DmarcRejectAction' },
            @{ Label = 'If the message is detected as spoof'; Prop = 'AuthenticationFailAction' },
            @{ Label = 'Quarantine policy (spoof)'; Prop = 'SpoofQuarantineTag' },
            @{ Label = 'Show first contact safety tip'; Prop = 'EnableFirstContactSafetyTips' },
            @{ Label = 'Show user impersonation safety tip'; Prop = 'EnableSimilarUsersSafetyTips' },
            @{ Label = 'Show domain impersonation safety tip'; Prop = 'EnableSimilarDomainsSafetyTips' },
            @{ Label = 'Show unusual characters safety tip'; Prop = 'EnableUnusualCharactersSafetyTips' },
            @{ Label = 'Show (?) for unauthenticated senders'; Prop = 'EnableUnauthenticatedSender' },
            @{ Label = 'Show "via" tag'; Prop = 'EnableViaTag' }
        )
        $phish = @(); $phishFailed = $false
        try { $phish = @(Get-AntiPhishPolicy) } catch { $phishFailed = $true; Add-LogEntry -Section 'AntiPhishPolicy' -ErrorRecord $_ }
        $phishRules = @(); $phishRulesFailed = $false
        try { $phishRules = @(Get-AntiPhishRule) } catch { $phishRulesFailed = $true; Add-LogEntry -Section 'AntiPhishRule' -ErrorRecord $_ }
        $phishJoined = @($phish | ForEach-Object {
            $pol = $_; $r = @($phishRules | Where-Object { $_.AntiPhishPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -RuleLookupFailed:($phishRulesFailed -or $eopPresetFailed)) }
        })
        $phishJoined = @($phishJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $pRows = @($phishJoined | ForEach-Object {
            $o = [ordered]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                             IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                             ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains }
            foreach ($s in $phishSettings) {
                $col = if ($o.Contains($s.Prop)) { "Policy$($s.Prop)" } else { $s.Prop }
                $o[$col] = $_.Policy.($s.Prop)
            }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'AntiPhishPolicies' -Description 'Anti-phishing policies' -Rows $pRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$phishFailed
        foreach ($j in $phishJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            $card = [ordered]@{
                'Status'           = $j.Info.Status
                'Priority'         = $j.Info.Priority
                'Applies to'       = $j.Info.AppliesTo
                'Included users'   = @{ Value = $j.Info.IncludedUsers; CsvFile = 'AntiPhishPolicies.csv' }
                'Included groups'  = @{ Value = $j.Info.IncludedGroups; CsvFile = 'AntiPhishPolicies.csv' }
                'Included domains' = @{ Value = $j.Info.IncludedDomains; CsvFile = 'AntiPhishPolicies.csv' }
                'Excluded users'   = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'AntiPhishPolicies.csv' }
                'Excluded groups'  = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'AntiPhishPolicies.csv' }
                'Excluded domains' = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'AntiPhishPolicies.csv' }
            }
            $phishActionProps = @('TargetedUserProtectionAction', 'TargetedDomainProtectionAction', 'MailboxIntelligenceProtectionAction', 'AuthenticationFailAction', 'DmarcQuarantineAction', 'DmarcRejectAction')
            foreach ($s in $phishSettings) {
                $v = $j.Policy.($s.Prop)
                if ($phishActionProps -contains $s.Prop) { $v = Format-ThreatAction $v }
                $card[$s.Label] = @{ Value = $v; CsvFile = 'AntiPhishPolicies.csv' }
            }
            Out-Card -Settings $card
        }

        Add-Line "## Anti-spam inbound policies"
        Add-Line
        $spamSettings = @(
            @{ Label = 'Bulk email threshold'; Prop = 'BulkThreshold' },
            @{ Label = 'Contains specific languages'; Prop = 'EnableLanguageBlockList' },
            @{ Label = 'Languages'; Prop = 'LanguageBlockList' },
            @{ Label = 'From these countries'; Prop = 'EnableRegionBlockList' },
            @{ Label = 'Countries'; Prop = 'RegionBlockList' },
            @{ Label = 'Spam'; Prop = 'SpamAction' },
            @{ Label = 'Quarantine policy (spam)'; Prop = 'SpamQuarantineTag' },
            @{ Label = 'High confidence spam'; Prop = 'HighConfidenceSpamAction' },
            @{ Label = 'Quarantine policy (high confidence spam)'; Prop = 'HighConfidenceSpamQuarantineTag' },
            @{ Label = 'Phishing'; Prop = 'PhishSpamAction' },
            @{ Label = 'Quarantine policy (phishing)'; Prop = 'PhishQuarantineTag' },
            @{ Label = 'High confidence phishing'; Prop = 'HighConfidencePhishAction' },
            @{ Label = 'Quarantine policy (high confidence phishing)'; Prop = 'HighConfidencePhishQuarantineTag' },
            @{ Label = 'Bulk compliant level (BCL) met or exceeded'; Prop = 'BulkSpamAction' },
            @{ Label = 'Quarantine policy (bulk)'; Prop = 'BulkQuarantineTag' },
            @{ Label = 'Intra-organizational messages to act on'; Prop = 'IntraOrgFilterState' },
            @{ Label = 'Retain spam in quarantine for this many days'; Prop = 'QuarantineRetentionPeriod' },
            @{ Label = 'Enable spam safety tips'; Prop = 'InlineSafetyTipsEnabled' },
            @{ Label = 'Enable zero-hour auto purge (ZAP) for phishing'; Prop = 'PhishZapEnabled' },
            @{ Label = 'Enable ZAP for spam'; Prop = 'SpamZapEnabled' },
            @{ Label = 'Allowed senders'; Prop = 'AllowedSenders' },
            @{ Label = 'Allowed domains'; Prop = 'AllowedSenderDomains' },
            @{ Label = 'Blocked senders'; Prop = 'BlockedSenders' },
            @{ Label = 'Blocked domains'; Prop = 'BlockedSenderDomains' }
        )
        $asfProps = @('IncreaseScoreWithImageLinks', 'IncreaseScoreWithNumericIps', 'IncreaseScoreWithRedirectToOtherPort',
                      'IncreaseScoreWithBizOrInfoUrls', 'MarkAsSpamEmptyMessages', 'MarkAsSpamJavaScriptInHtml',
                      'MarkAsSpamFramesInHtml', 'MarkAsSpamObjectTagsInHtml', 'MarkAsSpamEmbedTagsInHtml',
                      'MarkAsSpamFormTagsInHtml', 'MarkAsSpamWebBugsInHtml', 'MarkAsSpamSensitiveWordList',
                      'MarkAsSpamSpfRecordHardFail', 'MarkAsSpamFromAddressAuthFail', 'MarkAsSpamNdrBackscatter',
                      'MarkAsSpamBulkMail')
        $spamIn = @(); $spamInFailed = $false
        try { $spamIn = @(Get-HostedContentFilterPolicy) } catch { $spamInFailed = $true; Add-LogEntry -Section 'HostedContentFilterPolicy' -ErrorRecord $_ }
        $spamInRules = @()
        $spamInRulesFailed = $false
        try { $spamInRules = @(Get-HostedContentFilterRule) } catch { $spamInRulesFailed = $true; Add-LogEntry -Section 'HostedContentFilterRule' -ErrorRecord $_ }
        $siJoined = @($spamIn | ForEach-Object {
            $pol = $_; $r = @($spamInRules | Where-Object { $_.HostedContentFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -RuleLookupFailed:($spamInRulesFailed -or $eopPresetFailed)) }
        })
        $siJoined = @($siJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $siRows = @($siJoined | ForEach-Object {
            $o = [ordered]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                             IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                             ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains }
            foreach ($s in $spamSettings) { $o[$s.Prop] = $_.Policy.($s.Prop) }
            foreach ($a in $asfProps) { $o[$a] = $_.Policy.$a }
            $o.TestModeAction = $_.Policy.TestModeAction
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'AntiSpamInboundPolicies' -Description 'Anti-spam inbound policies' -Rows $siRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$spamInFailed
        foreach ($j in $siJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            $card = [ordered]@{
                'Status'           = $j.Info.Status
                'Priority'         = $j.Info.Priority
                'Applies to'       = $j.Info.AppliesTo
                'Included users'   = @{ Value = $j.Info.IncludedUsers; CsvFile = 'AntiSpamInboundPolicies.csv' }
                'Included groups'  = @{ Value = $j.Info.IncludedGroups; CsvFile = 'AntiSpamInboundPolicies.csv' }
                'Included domains' = @{ Value = $j.Info.IncludedDomains; CsvFile = 'AntiSpamInboundPolicies.csv' }
                'Excluded users'   = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'AntiSpamInboundPolicies.csv' }
                'Excluded groups'  = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'AntiSpamInboundPolicies.csv' }
                'Excluded domains' = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'AntiSpamInboundPolicies.csv' }
            }
            $spamActionProps = @('SpamAction', 'HighConfidenceSpamAction', 'PhishSpamAction', 'HighConfidencePhishAction', 'BulkSpamAction')
            foreach ($s in $spamSettings) {
                $v = $j.Policy.($s.Prop)
                if ($spamActionProps -contains $s.Prop) { $v = Format-ThreatAction $v }
                $card[$s.Label] = @{ Value = $v; CsvFile = 'AntiSpamInboundPolicies.csv' }
            }
            $asfNames = @{
                IncreaseScoreWithImageLinks = 'Image links to remote sites'; IncreaseScoreWithNumericIps = 'Numeric IP address in URL'
                IncreaseScoreWithRedirectToOtherPort = 'URL redirect to other port'; IncreaseScoreWithBizOrInfoUrls = 'Links to .biz or .info websites'
                MarkAsSpamEmptyMessages = 'Empty messages'; MarkAsSpamJavaScriptInHtml = 'JavaScript or VBScript in HTML'
                MarkAsSpamFramesInHtml = 'Frame or iframe tags in HTML'; MarkAsSpamObjectTagsInHtml = 'Object tags in HTML'
                MarkAsSpamEmbedTagsInHtml = 'Embed tags in HTML'; MarkAsSpamFormTagsInHtml = 'Form tags in HTML'
                MarkAsSpamWebBugsInHtml = 'Web bugs in HTML'; MarkAsSpamSensitiveWordList = 'Apply sensitive word list'
                MarkAsSpamSpfRecordHardFail = 'SPF record: hard fail'; MarkAsSpamFromAddressAuthFail = 'Sender ID filtering: hard fail'
                MarkAsSpamNdrBackscatter = 'Backscatter'; MarkAsSpamBulkMail = 'Mark bulk email as spam'
            }
            $asfOn = @($asfProps | Where-Object { $v = $j.Policy.$_; "$v" -match '\S' -and "$v" -ne 'Off' } | ForEach-Object {
                $n = $asfNames[$_]
                if ("$($j.Policy.$_)" -eq 'Test') { "$n (test)" }
                elseif ("$($j.Policy.$_)" -eq 'On') { $n }
                else { "$n ($($j.Policy.$_))" }
            })
            $card['Advanced spam filter (ASF) settings turned on'] = if ($asfOn.Count -gt 0) { $asfOn -join ', ' } else { 'None (all Off)' }
            $card['ASF test mode'] = $j.Policy.TestModeAction
            Out-Card -Settings $card
        }

        Add-Line "## Connection filter policy"
        Add-Line
        $cf = @(); $cfFailed = $false
        try { $cf = @(Get-HostedConnectionFilterPolicy) } catch { $cfFailed = $true; Add-LogEntry -Section 'HostedConnectionFilterPolicy' -ErrorRecord $_ }
        $cfRows = @($cf | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Status = 'On (tenant-wide policy)'; IPAllowList = $_.IPAllowList; IPBlockList = $_.IPBlockList; EnableSafeList = $_.EnableSafeList }
        })
        Out-Overview -CsvName 'ConnectionFilterPolicy' -Description 'Connection filter policies' -Rows $cfRows -Columns @('Name','Status','EnableSafeList') -Failed:$cfFailed
        foreach ($pol in $cf) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Status'                                                    = 'On (tenant-wide policy)'
                'Always allow messages from the following IP addresses'     = @{ Value = $pol.IPAllowList; CsvFile = 'ConnectionFilterPolicy.csv' }
                'Always block messages from the following IP addresses'     = @{ Value = $pol.IPBlockList; CsvFile = 'ConnectionFilterPolicy.csv' }
                'Turn on safe list'                                         = $pol.EnableSafeList
            })
        }

        Add-Line "## Outbound spam policies"
        Add-Line
        $spamOut = @(); $spamOutFailed = $false
        try { $spamOut = @(Get-HostedOutboundSpamFilterPolicy) } catch { $spamOutFailed = $true; Add-LogEntry -Section 'HostedOutboundSpamFilterPolicy' -ErrorRecord $_ }
        $spamOutRules = @()
        $spamOutRulesFailed = $false
        try { $spamOutRules = @(Get-HostedOutboundSpamFilterRule) } catch { $spamOutRulesFailed = $true; Add-LogEntry -Section 'HostedOutboundSpamFilterRule' -ErrorRecord $_ }
        $soJoined = @($spamOut | ForEach-Object {
            $pol = $_; $r = @($spamOutRules | Where-Object { $_.HostedOutboundSpamFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -SenderBased -RuleLookupFailed:$spamOutRulesFailed) }
        })
        $soJoined = @($soJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $soRows = @($soJoined | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                              IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                              ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains
                              RecipientLimitExternalPerHour = $_.Policy.RecipientLimitExternalPerHour
                              RecipientLimitInternalPerHour = $_.Policy.RecipientLimitInternalPerHour; RecipientLimitPerDay = $_.Policy.RecipientLimitPerDay
                              ActionWhenThresholdReached = $_.Policy.ActionWhenThresholdReached
                              AutoForwardingMode = $_.Policy.AutoForwardingMode; BccSuspiciousOutboundMail = $_.Policy.BccSuspiciousOutboundMail
                              BccSuspiciousOutboundAdditionalRecipients = $_.Policy.BccSuspiciousOutboundAdditionalRecipients
                              NotifyOutboundSpam = $_.Policy.NotifyOutboundSpam
                              NotifyOutboundSpamRecipients = $_.Policy.NotifyOutboundSpamRecipients }
        })
        Out-Overview -CsvName 'OutboundSpamPolicies' -Description 'Outbound spam policies' -Rows $soRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$spamOutFailed
        foreach ($j in $soJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Status'                            = $j.Info.Status
                'Priority'                          = $j.Info.Priority
                'Applies to'                        = $j.Info.AppliesTo
                'Included senders'                  = @{ Value = $j.Info.IncludedUsers; CsvFile = 'OutboundSpamPolicies.csv' }
                'Included sender groups'            = @{ Value = $j.Info.IncludedGroups; CsvFile = 'OutboundSpamPolicies.csv' }
                'Included sender domains'           = @{ Value = $j.Info.IncludedDomains; CsvFile = 'OutboundSpamPolicies.csv' }
                'Excluded senders'                  = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'OutboundSpamPolicies.csv' }
                'Excluded sender groups'            = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'OutboundSpamPolicies.csv' }
                'Excluded sender domains'           = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'OutboundSpamPolicies.csv' }
                'External message limit per hour'                          = $(if ("$($j.Policy.RecipientLimitExternalPerHour)" -eq '0') { '0 (service default)' } else { $j.Policy.RecipientLimitExternalPerHour })
                'Internal message limit per hour'                          = $(if ("$($j.Policy.RecipientLimitInternalPerHour)" -eq '0') { '0 (service default)' } else { $j.Policy.RecipientLimitInternalPerHour })
                'Daily message limit'                                      = $(if ("$($j.Policy.RecipientLimitPerDay)" -eq '0') { '0 (service default)' } else { $j.Policy.RecipientLimitPerDay })
                'Restriction placed on users who reach the message limit'  = $j.Policy.ActionWhenThresholdReached
                'Automatic forwarding rules'                               = $j.Policy.AutoForwardingMode
                'Send a copy of suspicious outbound messages'              = $j.Policy.BccSuspiciousOutboundMail
                'Send a copy of suspicious outbound messages to'           = @{ Value = $j.Policy.BccSuspiciousOutboundAdditionalRecipients; CsvFile = 'OutboundSpamPolicies.csv' }
                'Notify these users when a sender is blocked'              = $j.Policy.NotifyOutboundSpam
                'Blocked-sender notification recipients'                   = @{ Value = $j.Policy.NotifyOutboundSpamRecipients; CsvFile = 'OutboundSpamPolicies.csv' }
            })
        }

        Add-Line "## Anti-malware policies"
        Add-Line
        $mal = @(); $malFailed = $false
        try { $mal = @(Get-MalwareFilterPolicy) } catch { $malFailed = $true; Add-LogEntry -Section 'MalwareFilterPolicy' -ErrorRecord $_ }
        $malRules = @()
        $malRulesFailed = $false
        try { $malRules = @(Get-MalwareFilterRule) } catch { $malRulesFailed = $true; Add-LogEntry -Section 'MalwareFilterRule' -ErrorRecord $_ }
        $malJoined = @($mal | ForEach-Object {
            $pol = $_; $r = @($malRules | Where-Object { $_.MalwareFilterPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -RuleLookupFailed:($malRulesFailed -or $eopPresetFailed)) }
        })
        $malJoined = @($malJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $malRows = @($malJoined | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                              IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                              ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains
                              EnableFileFilter = $_.Policy.EnableFileFilter; FileTypes = $_.Policy.FileTypes; FileTypeAction = $_.Policy.FileTypeAction
                              ZapEnabled = $_.Policy.ZapEnabled; QuarantineTag = $_.Policy.QuarantineTag
                              EnableInternalSenderAdminNotifications = $_.Policy.EnableInternalSenderAdminNotifications
                              InternalSenderAdminAddress = $_.Policy.InternalSenderAdminAddress
                              EnableExternalSenderAdminNotifications = $_.Policy.EnableExternalSenderAdminNotifications
                              ExternalSenderAdminAddress = $_.Policy.ExternalSenderAdminAddress
                              CustomNotifications = $_.Policy.CustomNotifications; CustomFromName = $_.Policy.CustomFromName; CustomFromAddress = $_.Policy.CustomFromAddress }
        })
        Out-Overview -CsvName 'MalwareFilterPolicies' -Description 'Anti-malware policies' -Rows $malRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$malFailed
        foreach ($j in $malJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Status'                                   = $j.Info.Status
                'Priority'                                 = $j.Info.Priority
                'Applies to'                               = $j.Info.AppliesTo
                'Included users'                           = @{ Value = $j.Info.IncludedUsers; CsvFile = 'MalwareFilterPolicies.csv' }
                'Included groups'                          = @{ Value = $j.Info.IncludedGroups; CsvFile = 'MalwareFilterPolicies.csv' }
                'Included domains'                         = @{ Value = $j.Info.IncludedDomains; CsvFile = 'MalwareFilterPolicies.csv' }
                'Excluded users'                           = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'MalwareFilterPolicies.csv' }
                'Excluded groups'                          = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'MalwareFilterPolicies.csv' }
                'Excluded domains'                         = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'MalwareFilterPolicies.csv' }
                'Enable the common attachments filter'     = $j.Policy.EnableFileFilter
                'File types'                               = @{ Value = $j.Policy.FileTypes; CsvFile = 'MalwareFilterPolicies.csv' }
                'When these file types are found'          = $(if ("$($j.Policy.FileTypeAction)" -eq 'Reject') { 'Reject the message with a non-delivery report (NDR)' } else { Format-ThreatAction $j.Policy.FileTypeAction })
                'Enable zero-hour auto purge for malware'  = $j.Policy.ZapEnabled
                'Quarantine policy'                        = $j.Policy.QuarantineTag
                'Notify an admin about undelivered messages from internal senders' = $j.Policy.EnableInternalSenderAdminNotifications
                'Internal admin email address'             = $j.Policy.InternalSenderAdminAddress
                'Notify an admin about undelivered messages from external senders' = $j.Policy.EnableExternalSenderAdminNotifications
                'External admin email address'             = $j.Policy.ExternalSenderAdminAddress
                'Customize notifications'                  = $j.Policy.CustomNotifications
                'From name'                                = $j.Policy.CustomFromName
                'From address'                             = $j.Policy.CustomFromAddress
            })
        }

        Add-Line "## Safe Attachments policies"
        Add-Line
        $sa = @(); $saFailed = $false
        try { $sa = @(Get-SafeAttachmentPolicy) } catch { $saFailed = $true; Add-LogEntry -Section 'SafeAttachmentPolicy' -ErrorRecord $_ }
        $saRules = @()
        $saRulesFailed = $false
        try { $saRules = @(Get-SafeAttachmentRule) } catch { $saRulesFailed = $true; Add-LogEntry -Section 'SafeAttachmentRule' -ErrorRecord $_ }
        $saJoined = @($sa | ForEach-Object {
            $pol = $_; $r = @($saRules | Where-Object { $_.SafeAttachmentPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -RuleLookupFailed:($saRulesFailed -or $atpPresetFailed -or $builtinFailed)) }
        })
        $saJoined = @($saJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $saRows = @($saJoined | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                              IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                              ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains
                              Enable = $_.Policy.Enable; Action = $_.Policy.Action; QuarantineTag = $_.Policy.QuarantineTag
                              Redirect = $_.Policy.Redirect; RedirectAddress = $_.Policy.RedirectAddress
                              EnableBlockingEncryptedAttachments = $_.Policy.EnableBlockingEncryptedAttachments
                              ExcludedTypesFromBlockingEncryptedAttachments = $_.Policy.ExcludedTypesFromBlockingEncryptedAttachments
                              QuarantineTagForBlockingEncryptedAttachments = $_.Policy.QuarantineTagForBlockingEncryptedAttachments }
        })
        Out-Overview -CsvName 'SafeAttachmentPolicies' -Description 'Safe Attachments policies' -Rows $saRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$saFailed
        foreach ($j in $saJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            $saResponse = if ($j.Policy.Enable -eq $false) { 'Off' }
                else { switch ("$($j.Policy.Action)") { 'Allow' { 'Monitor' } 'Block' { 'Block' } 'DynamicDelivery' { 'Dynamic Delivery' } default { "$($j.Policy.Action)" } } }
            Out-Card -Settings ([ordered]@{
                'Status'           = $j.Info.Status
                'Priority'         = $j.Info.Priority
                'Applies to'       = $j.Info.AppliesTo
                'Included users'   = @{ Value = $j.Info.IncludedUsers; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Included groups'  = @{ Value = $j.Info.IncludedGroups; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Included domains' = @{ Value = $j.Info.IncludedDomains; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Excluded users'   = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Excluded groups'  = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Excluded domains' = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Safe Attachments unknown malware response' = $saResponse
                'Quarantine policy' = $j.Policy.QuarantineTag
                'Redirect messages with detected attachments' = $j.Policy.Redirect
                'Redirect address' = $j.Policy.RedirectAddress
                'Block unscanned attachments' = $j.Policy.EnableBlockingEncryptedAttachments
                'Excluded file types' = @{ Value = $j.Policy.ExcludedTypesFromBlockingEncryptedAttachments; CsvFile = 'SafeAttachmentPolicies.csv' }
                'Quarantine policy (unscanned attachments)' = $j.Policy.QuarantineTagForBlockingEncryptedAttachments
            })
        }

        Add-Line "## ATP policy for Office 365"
        Add-Line
        $atp = @(); $atpFailed = $false
        try { $atp = @(Get-AtpPolicyForO365) } catch { $atpFailed = $true; Add-LogEntry -Section 'AtpPolicyForO365' -ErrorRecord $_ }
        $atpRows = @($atp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; EnableATPForSPOTeamsODB = $_.EnableATPForSPOTeamsODB; EnableSafeDocs = $_.EnableSafeDocs
                              AllowSafeDocsOpen = $_.AllowSafeDocsOpen }
        })
        Out-Overview -CsvName 'AtpPolicyForO365' -Description 'ATP policy for Office 365' -Rows $atpRows -Columns @('Name','EnableATPForSPOTeamsODB','EnableSafeDocs') -Failed:$atpFailed
        foreach ($pol in $atp) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Turn on Defender for Office 365 for SharePoint, OneDrive, and Microsoft Teams' = $pol.EnableATPForSPOTeamsODB
                'Turn on Safe Documents for Office clients' = $pol.EnableSafeDocs
                'Allow people to click through Protected View even if Safe Documents identified the file as malicious' = $pol.AllowSafeDocsOpen
            })
        }

        Add-Line "## Safe Links policies"
        Add-Line
        $slSettings = @(
            @{ Label = 'On: Safe Links checks a list of known malicious links when users click links in email'; Prop = 'EnableSafeLinksForEmail' },
            @{ Label = 'Apply Safe Links to email messages sent within the organization'; Prop = 'EnableForInternalSenders' },
            @{ Label = 'Apply real-time URL scanning for suspicious links and links that point to files'; Prop = 'ScanUrls' },
            @{ Label = 'Wait for URL scanning to complete before delivering the message'; Prop = 'DeliverMessageAfterScan' },
            @{ Label = 'Do not rewrite URLs, do checks via Safe Links API only'; Prop = 'DisableUrlRewrite' },
            @{ Label = 'Do not rewrite the following URLs in email'; Prop = 'DoNotRewriteUrls' },
            @{ Label = 'Teams'; Prop = 'EnableSafeLinksForTeams' },
            @{ Label = 'Office 365 apps'; Prop = 'EnableSafeLinksForOffice' },
            @{ Label = 'Track user clicks'; Prop = 'TrackClicks' },
            @{ Label = 'Let users click through to the original URL'; Prop = 'AllowClickThrough' },
            @{ Label = 'Display the organization branding on notification and warning pages'; Prop = 'EnableOrganizationBranding' },
            @{ Label = 'Use custom notification text'; Prop = 'CustomNotificationText' },
            @{ Label = 'Use Microsoft Translator for automatic localization'; Prop = 'UseTranslatedNotificationText' }
        )
        $sl = @(); $slFailed = $false
        try { $sl = @(Get-SafeLinksPolicy) } catch { $slFailed = $true; Add-LogEntry -Section 'SafeLinksPolicy' -ErrorRecord $_ }
        $slRules = @()
        $slRulesFailed = $false
        try { $slRules = @(Get-SafeLinksRule) } catch { $slRulesFailed = $true; Add-LogEntry -Section 'SafeLinksRule' -ErrorRecord $_ }
        $slJoined = @($sl | ForEach-Object {
            $pol = $_; $r = @($slRules | Where-Object { $_.SafeLinksPolicy -eq $pol.Name }) | Select-Object -First 1
            [PSCustomObject]@{ Policy = $pol; Info = (Get-PolicyRuleInfo -Policy $pol -Rule $r -PresetRules $presetRules -BuiltInRules $builtin -RuleLookupFailed:($slRulesFailed -or $atpPresetFailed -or $builtinFailed)) }
        })
        $slJoined = @($slJoined | Sort-Object { $_.Info.SortKey }, { "$($_.Policy.Name)" })
        $slRows = @($slJoined | ForEach-Object {
            $o = [ordered]@{ Name = $_.Policy.Name; Status = $_.Info.Status; Priority = $_.Info.Priority; AppliesTo = $_.Info.AppliesTo
                             IncludedUsers = $_.Info.IncludedUsers; IncludedGroups = $_.Info.IncludedGroups; IncludedDomains = $_.Info.IncludedDomains
                             ExcludedUsers = $_.Info.ExcludedUsers; ExcludedGroups = $_.Info.ExcludedGroups; ExcludedDomains = $_.Info.ExcludedDomains }
            foreach ($s in $slSettings) { $o[$s.Prop] = $_.Policy.($s.Prop) }
            [PSCustomObject]$o
        })
        Out-Overview -CsvName 'SafeLinksPolicies' -Description 'Safe Links policies' -Rows $slRows -Columns @('Name','Status','Priority','AppliesTo') -Failed:$slFailed
        foreach ($j in $slJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            $card = [ordered]@{
                'Status'           = $j.Info.Status
                'Priority'         = $j.Info.Priority
                'Applies to'       = $j.Info.AppliesTo
                'Included users'   = @{ Value = $j.Info.IncludedUsers; CsvFile = 'SafeLinksPolicies.csv' }
                'Included groups'  = @{ Value = $j.Info.IncludedGroups; CsvFile = 'SafeLinksPolicies.csv' }
                'Included domains' = @{ Value = $j.Info.IncludedDomains; CsvFile = 'SafeLinksPolicies.csv' }
                'Excluded users'   = @{ Value = $j.Info.ExcludedUsers; CsvFile = 'SafeLinksPolicies.csv' }
                'Excluded groups'  = @{ Value = $j.Info.ExcludedGroups; CsvFile = 'SafeLinksPolicies.csv' }
                'Excluded domains' = @{ Value = $j.Info.ExcludedDomains; CsvFile = 'SafeLinksPolicies.csv' }
            }
            foreach ($s in $slSettings) { $card[$s.Label] = @{ Value = $j.Policy.($s.Prop); CsvFile = 'SafeLinksPolicies.csv' } }
            Out-Card -Settings $card
        }

        Add-Line "## Quarantine policies"
        Add-Line
        Add-Line "Quarantine policies have no on/off state. A quarantine policy takes effect when a threat policy (anti-spam, anti-phishing, anti-malware, Safe Attachments) references it as its quarantine policy."
        Add-Line
        $qp = @(); $qpFailed = $false
        try { $qp = @(Get-QuarantinePolicy) } catch { $qpFailed = $true; Add-LogEntry -Section 'QuarantinePolicy' -ErrorRecord $_ }
        $qpJoined = @($qp | ForEach-Object { [PSCustomObject]@{ Policy = $_; Info = (Get-QuarantinePermissionInfo -Policy $_) } })
        $qpRows = @($qpJoined | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Policy.Name; RecipientMessageAccess = $_.Info.Access; Permissions = $_.Info.Permissions
                              QuarantineNotification = $_.Info.Notifications
                              IncludeMessagesFromBlockedSenderAddress = $_.Policy.IncludeMessagesFromBlockedSenderAddress
                              EndUserQuarantinePermissionsValue = $_.Policy.EndUserQuarantinePermissionsValue; ESNEnabled = $_.Policy.ESNEnabled }
        })
        Out-Overview -CsvName 'QuarantinePolicies' -Description 'Quarantine policies' -Rows $qpRows -Columns @('Name','RecipientMessageAccess','QuarantineNotification') -Failed:$qpFailed
        foreach ($j in $qpJoined) {
            Add-Line "### $(ConvertTo-MdHeading $j.Policy.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Recipient message access'                        = $j.Info.Access
                'User message access'                             = $j.Info.Permissions
                'Quarantine notification'                         = $j.Info.Notifications
                'Include quarantined messages from blocked sender addresses' = $j.Policy.IncludeMessagesFromBlockedSenderAddress
            })
        }

        Add-Line "### Global quarantine notification settings"
        Add-Line
        $gq = @(); $gqFailed = $false
        try { $gq = @(Get-QuarantinePolicy -QuarantinePolicyType GlobalQuarantinePolicy) } catch { $gqFailed = $true; Add-LogEntry -Section 'GlobalQuarantinePolicy' -ErrorRecord $_ }
        if ($gq.Count -gt 0) {
            $g = $gq[0]
            $freq = $g.EndUserSpamNotificationFrequency
            [timespan]$ts = [timespan]::Zero
            $parsed = $false
            if ($freq -is [timespan]) { $ts = $freq; $parsed = $true }
            elseif ([timespan]::TryParse("$freq", [ref]$ts)) { $parsed = $true }
            if ($parsed) {
                if ($ts.TotalDays -ge 1) { $freq = if ([int]$ts.TotalDays -eq 1) { '1 day' } else { "$([int]$ts.TotalDays) days" } }
                else { $freq = "$([int]$ts.TotalHours) hours" }
            }
            Out-Card -Settings ([ordered]@{
                'Send end-user spam notifications every' = $freq
                'Custom sender address'                  = $g.EndUserSpamNotificationCustomFromAddress
                'Use my company logo'                    = $g.OrganizationBrandingEnabled
                'Languages / display names'              = @{ Value = $g.MultiLanguageSenderName; CsvFile = 'QuarantineGlobalSettings.csv' }
                'Disclaimer'                             = @{ Value = $g.MultiLanguageCustomDisclaimer; CsvFile = 'QuarantineGlobalSettings.csv' }
            })
        }
        elseif ($gqFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not available._"; Add-Line }
        Register-Csv -Name 'QuarantineGlobalSettings' -Description 'Global quarantine notification settings' -Rows $gq -Failed:$gqFailed

        Add-Line "## Tenant Allow/Block List"
        Add-Line
        $tablLists = [ordered]@{ 'Sender' = 'Domains & addresses'; 'Url' = 'URLs'; 'FileHash' = 'Files'; 'IP' = 'IP addresses' }
        $tablRows = @(); $tablFailed = @{}
        foreach ($listType in $tablLists.Keys) {
            try {
                $tablRows += @(Get-TenantAllowBlockListItems -ListType $listType -ErrorAction Stop | ForEach-Object {
                    [PSCustomObject]@{ List = $tablLists[$listType]; Value = $_.Value; Action = $_.Action; ExpirationDate = $(if ($_.ExpirationDate) { $_.ExpirationDate } else { 'Never expire' }); Notes = $_.Notes }
                })
            }
            catch { $tablFailed[$tablLists[$listType]] = $true; Add-LogEntry -Section "TenantAllowBlockListItems ($listType)" -ErrorRecord $_ }
        }
        Add-Line "Entries per list (full lists in ``TenantAllowBlockList.csv``):"
        Add-Line
        $tablCard = [ordered]@{}
        foreach ($name in $tablLists.Values) { $tablCard[$name] = $(if ($tablFailed[$name]) { $script:LookupFailedText } else { @($tablRows | Where-Object { $_.List -eq $name }).Count }) }
        Out-Card -Settings $tablCard
        foreach ($name in $tablLists.Values) {
            Add-Line "### $(ConvertTo-MdHeading $name)"
            Add-Line
            $listRows = @($tablRows | Where-Object { $_.List -eq $name } | Select-Object Value, Action, ExpirationDate, Notes)
            if ($tablFailed[$name]) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line; continue }
            if ($listRows.Count -eq 0) { Add-Line "_No entries._"; Add-Line; continue }
            Out-Overview -CsvName 'TenantAllowBlockList' -Description 'Tenant Allow/Block List entries (all lists)' -Rows $listRows -Columns @('Value','Action','ExpirationDate','Notes') -SkipCsv
        }
        $tablFailedNames = @($tablLists.Values | Where-Object { $tablFailed[$_] })
        if ($tablFailedNames.Count -ge $tablLists.Count) {
            Register-Csv -Name 'TenantAllowBlockList' -Description 'Tenant Allow/Block List entries (all lists)' -Rows @() -Failed
        }
        else {
            Register-Csv -Name 'TenantAllowBlockList' -Description 'Tenant Allow/Block List entries (all lists)' -Rows $tablRows -Note $(if ($tablFailedNames.Count) { "incomplete: $($tablFailedNames -join ', ') lookup failed" })
        }
        Add-Line "### Spoofed senders"
        Add-Line
        $spoof = @(); $spoofFailed = $false
        try { $spoof = @(Get-TenantAllowBlockListSpoofItems) } catch { $spoofFailed = $true; Add-LogEntry -Section 'TenantAllowBlockListSpoofItems' -ErrorRecord $_ }
        $spoofRows = @($spoof | ForEach-Object {
            [PSCustomObject]@{ SpoofedUser = $_.SpoofedUser; SendingInfrastructure = $_.SendingInfrastructure; Action = $_.Action; SpoofType = $_.SpoofType }
        })
        Out-Overview -CsvName 'TenantAllowBlockListSpoof' -Description 'Spoof intelligence entries' -Rows $spoofRows -Columns @('SpoofedUser','SendingInfrastructure','Action','SpoofType') -Failed:$spoofFailed

        Add-Line "## Advanced delivery"
        Add-Line
        $secOps = @(); $secOpsFailed = $false
        try { $secOps = @(Get-SecOpsOverridePolicy) } catch { $secOpsFailed = $true; Add-LogEntry -Section 'SecOpsOverridePolicy' -ErrorRecord $_ }
        $secOpsRules = @()
        $secOpsRulesFailed = $false
        try { $secOpsRules = @(Get-ExoSecOpsOverrideRule) } catch { $secOpsRulesFailed = $true; Add-LogEntry -Section 'ExoSecOpsOverrideRule' -ErrorRecord $_ }
        $secOpsRows = @($secOps | ForEach-Object {
            $pol = $_
            $ids = @("$($pol.Identity)", "$($pol.Name)", "$($pol.Guid)", "$($pol.ExchangeObjectId)", "$($pol.Id)", "$($pol.DistinguishedName)") | Where-Object { $_ }
            $r = @($secOpsRules | Where-Object { "$($_.Policy)" -in $ids }) | Select-Object -First 1
            if (-not $r -and $secOps.Count -eq 1 -and $secOpsRules.Count -eq 1) { $r = $secOpsRules[0] }
            [PSCustomObject]@{ Name = $pol.Name; SentTo = $(if ($secOpsRulesFailed) { 'Unknown (rule lookup failed)' } else { $r.SentTo }); SentToMemberOf = $r.SentToMemberOf }
        })
        Add-Line "### SecOps mailboxes"
        Add-Line
        Out-Overview -CsvName 'SecOpsOverridePolicy' -Description 'SecOps override (mailbox) policies' -Rows $secOpsRows -Columns @('Name','SentTo','SentToMemberOf') -Failed:$secOpsFailed
        Add-Line "### Phishing simulation"
        Add-Line
        $phishSim = @(); $phishSimFailed = $false
        try { $phishSim = @(Get-PhishSimOverridePolicy) } catch { $phishSimFailed = $true; Add-LogEntry -Section 'PhishSimOverridePolicy' -ErrorRecord $_ }
        $phishSimRules = @()
        $phishSimRulesFailed = $false
        try { $phishSimRules = @(Get-ExoPhishSimOverrideRule) } catch { $phishSimRulesFailed = $true; Add-LogEntry -Section 'ExoPhishSimOverrideRule' -ErrorRecord $_ }
        $simRows = @($phishSim | ForEach-Object {
            $pol = $_
            $ids = @("$($pol.Identity)", "$($pol.Name)", "$($pol.Guid)", "$($pol.ExchangeObjectId)", "$($pol.Id)", "$($pol.DistinguishedName)") | Where-Object { $_ }
            $r = @($phishSimRules | Where-Object { "$($_.Policy)" -in $ids }) | Select-Object -First 1
            if (-not $r -and $phishSim.Count -eq 1 -and $phishSimRules.Count -eq 1) { $r = $phishSimRules[0] }
            [PSCustomObject]@{ Name = $pol.Name; Domains = $(if ($phishSimRulesFailed) { 'Unknown (rule lookup failed)' } else { $r.Domains }); SenderIpRanges = $r.SenderIpRanges }
        })
        Out-Overview -CsvName 'PhishSimOverridePolicy' -Description 'Phishing simulation override policies' -Rows $simRows -Columns @('Name','Domains','SenderIpRanges') -Failed:$phishSimFailed
        Add-Line "### Simulation URLs to allow"
        Add-Line
        $simUrls = @(); $simUrlsFailed = $false
        try { $simUrls = @(Get-TenantAllowBlockListItems -ListType Url -ListSubType AdvancedDelivery) } catch { $simUrlsFailed = $true; Add-LogEntry -Section 'Simulation URLs' -ErrorRecord $_ }
        $simUrlRows = @($simUrls | ForEach-Object {
            [PSCustomObject]@{ SimulationUrl = $_.Value }
        })
        Out-Overview -CsvName 'SimulationUrls' -Description 'Simulation URLs' -Rows $simUrlRows -Columns @('SimulationUrl') -Failed:$simUrlsFailed

        Add-Line "## User-reported settings"
        Add-Line
        $rsp = @(); $rspFailed = $false
        try { $rsp = @(Get-ReportSubmissionPolicy) } catch { $rspFailed = $true; Add-LogEntry -Section 'ReportSubmissionPolicy' -ErrorRecord $_ }
        $rspRows = @($rsp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; DisableUserSubmissionOptions = $_.DisableUserSubmissionOptions
                              EnableReportToMicrosoft = $_.EnableReportToMicrosoft; ReportPhishToCustomizedAddress = $_.ReportPhishToCustomizedAddress
                              ReportJunkAddresses = $_.ReportJunkAddresses; ReportPhishAddresses = $_.ReportPhishAddresses
                              EnableThirdPartyAddress = $_.EnableThirdPartyAddress; ThirdPartyReportAddresses = $_.ThirdPartyReportAddresses
                              ReportChatMessageEnabled = $_.ReportChatMessageEnabled; PreSubmitMessageEnabled = $_.PreSubmitMessageEnabled
                              PostSubmitMessageEnabled = $_.PostSubmitMessageEnabled; EnableUserEmailNotification = $_.EnableUserEmailNotification
                              EnableCustomNotificationSender = $_.EnableCustomNotificationSender; NotificationSenderAddress = $_.NotificationSenderAddress
                              DisableQuarantineReportingOption = $_.DisableQuarantineReportingOption }
        })
        Out-Overview -CsvName 'ReportSubmissionPolicy' -Description 'User-reported message settings' -Rows $rspRows -Columns @('Name','EnableReportToMicrosoft','ReportPhishToCustomizedAddress') -Failed:$rspFailed
        foreach ($pol in $rsp) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            $sendTo = if ((Test-IsTrue $pol.EnableReportToMicrosoft) -and (Test-IsTrue $pol.ReportPhishToCustomizedAddress)) { 'Microsoft and my reporting mailbox' }
                      elseif (Test-IsTrue $pol.EnableReportToMicrosoft) { 'Microsoft only' }
                      elseif (Test-IsTrue $pol.ReportPhishToCustomizedAddress) { 'My reporting mailbox only' }
                      else { 'Not configured' }
            $reportBoxes = @(@($pol.ReportPhishAddresses) + @($pol.ReportJunkAddresses) | Where-Object { $_ } | Select-Object -Unique)
            Out-Card -Settings ([ordered]@{
                'Monitor reported messages in Outlook'   = $(if (Test-IsTrue $pol.DisableUserSubmissionOptions) { 'Off' } else { 'On' })
                'Send reported messages to'              = $sendTo
                'Reporting mailbox'                      = @{ Value = $reportBoxes; CsvFile = 'ReportSubmissionPolicy.csv' }
                'Third-party reporting tool'             = $pol.EnableThirdPartyAddress
                'Third-party addresses'                  = @{ Value = $pol.ThirdPartyReportAddresses; CsvFile = 'ReportSubmissionPolicy.csv' }
                'Report Teams chat messages'             = $pol.ReportChatMessageEnabled
                'Ask users to confirm before reporting'  = $pol.PreSubmitMessageEnabled
                'Show success message after reporting'   = $pol.PostSubmitMessageEnabled
                'Email notifications to users about admin review results' = $pol.EnableUserEmailNotification
                'Custom notification sender'             = $(if (Test-IsTrue $pol.EnableCustomNotificationSender) { $pol.NotificationSenderAddress } else { $null })
                'Report from quarantine allowed'         = $(if ($null -eq $pol.DisableQuarantineReportingOption) { $null } else { -not (Test-IsTrue $pol.DisableQuarantineReportingOption) })
            })
        }

        Add-Line "## Priority account protection"
        Add-Line
        $ets = @(); $etsFailed = $false
        try { $ets = @(Get-EmailTenantSettings) } catch { $etsFailed = $true; Add-LogEntry -Section 'EmailTenantSettings' -ErrorRecord $_ }
        $etsRows = @($ets | ForEach-Object {
            [PSCustomObject]@{ Identity = $_.Identity; EnablePriorityAccountProtection = $_.EnablePriorityAccountProtection }
        })
        Register-Csv -Name 'EmailTenantSettings' -Description 'Email tenant settings' -Rows $etsRows -Failed:$etsFailed
        Out-Card -Settings ([ordered]@{
            'Priority account protection' = $(if ($etsFailed) { $script:LookupFailedText } else { $etsRows[0].EnablePriorityAccountProtection })
        })
        $vips = @(); $vipsFailed = $false
        try { $vips = @(Get-User -IsVIP -ResultSize Unlimited) } catch { $vipsFailed = $true; Add-LogEntry -Section 'Priority accounts (Get-User -IsVIP)' -ErrorRecord $_ }
        $vipRows = @($vips | ForEach-Object {
            [PSCustomObject]@{ DisplayName = $_.DisplayName; UserPrincipalName = $_.UserPrincipalName }
        })
        Out-Card -Settings ([ordered]@{ 'Priority accounts' = $(if ($vipsFailed) { $script:LookupFailedText } else { $vips.Count }) })
        Out-Overview -CsvName 'PriorityAccounts' -Description 'Priority accounts (VIP users)' -Rows $vipRows -Columns @('DisplayName','UserPrincipalName') -Failed:$vipsFailed

        Add-Line "## Microsoft Teams protection"
        Add-Line
        $tpp = @(); $tppFailed = $false
        try { $tpp = @(Get-TeamsProtectionPolicy) } catch { $tppFailed = $true; Add-LogEntry -Section 'TeamsProtectionPolicy' -ErrorRecord $_ }
        $tppRows = @($tpp | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ZapEnabled = $_.ZapEnabled; HighConfidencePhishQuarantineTag = $_.HighConfidencePhishQuarantineTag; MalwareQuarantineTag = $_.MalwareQuarantineTag }
        })
        Register-Csv -Name 'TeamsProtectionPolicy' -Description 'Teams protection policy' -Rows $tppRows -Failed:$tppFailed
        if ($tppFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        elseif ($tpp.Count -eq 0) { Add-Line "_None found._"; Add-Line }
        foreach ($pol in $tpp) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Zero-hour auto purge (ZAP) for Teams' = $pol.ZapEnabled
                'Quarantine policy (high confidence phishing)' = $pol.HighConfidencePhishQuarantineTag
                'Quarantine policy (malware)' = $pol.MalwareQuarantineTag
            })
        }
    }
}

function Write-PermissionsSection {
    Invoke-Section -Title 'Permissions (RBAC)' -Body {
        Add-Line "Administrative and end-user permissions: role groups and their members, role assignment policies, custom roles, direct assignments, scopes and RBAC for Applications."
        Add-Line

        Add-Line "## Role groups and members"
        Add-Line
        $rg = @(); $rgFailed = $false
        try { $rg = @(Get-RoleGroup) } catch { $rgFailed = $true; Add-LogEntry -Section 'RoleGroup' -ErrorRecord $_ }
        $rgRows = @()
        foreach ($g in $rg) {
            $members = @()
            $membersFailed = $false
            try { $members = @(Get-RoleGroupMember -Identity $g.Name) } catch { $membersFailed = $true; Add-LogEntry -Section "RoleGroupMember ($($g.Name))" -ErrorRecord $_ }
            $info = Get-RoleGroupMemberInfo -Members $members
            $roles = @($g.Roles | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.Name) { "$($_.Name)" } else { "$_" } } | Sort-Object)
            $rgRows += [PSCustomObject]@{
                RoleGroup = (Get-RoleGroupDisplayName $g.Name); Members = $(if ($membersFailed) { 'Unknown (member lookup failed)' } else { $info.Summary }); UserMembers = $info.Users; GroupMembers = $info.Groups
                ManagedBy = $g.ManagedBy; Roles = $roles; MembersFailed = $membersFailed
            }
        }
        Out-Overview -CsvName 'RoleGroups' -Description 'Role groups and members' -Rows @($rgRows | ForEach-Object {
            [PSCustomObject]@{ RoleGroup = $_.RoleGroup; Members = $_.Members; UserMembers = ($_.UserMembers -join '; '); GroupMembers = ($_.GroupMembers -join '; '); ManagedBy = ($_.ManagedBy -join '; '); Roles = ($_.Roles -join '; ') }
        }) -Columns @('RoleGroup','Members') -Failed:$rgFailed
        foreach ($g in @($rgRows | Where-Object { $_.Members -ne 'No members' })) {
            Add-Line "### $(ConvertTo-MdHeading $g.RoleGroup)"
            Add-Line
            $card = [ordered]@{}
            if ($g.MembersFailed) { $card['Members'] = 'Unknown (member lookup failed)' }
            if ($g.Roles.Count) { $card['Roles'] = ($g.Roles -join ', ') }
            if ($g.GroupMembers.Count) { $card['Assigned groups'] = @{ Value = $g.GroupMembers; CsvFile = 'RoleGroups.csv' } }
            if ($g.UserMembers.Count)  { $card['Assigned users']  = @{ Value = $g.UserMembers; CsvFile = 'RoleGroups.csv' } }
            Out-Card -Settings $card
        }

        Add-Line "## Role assignment policies (user roles)"
        Add-Line
        $rap = @(); $rapFailed = $false
        try { $rap = @(Get-RoleAssignmentPolicy) } catch { $rapFailed = $true; Add-LogEntry -Section 'RoleAssignmentPolicy' -ErrorRecord $_ }
        $mbx = @(); $mbxFailed = $false
        try { $mbx = @(Get-SharedMailboxes) } catch { $mbxFailed = $true }
        $rapRows = @($rap | ForEach-Object {
            $rapName = $_.Name
            [PSCustomObject]@{
                Name = $rapName; IsDefault = $_.IsDefault; Description = $_.Description
                AssignedRoles = @($_.AssignedRoles | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.Name) { "$($_.Name)" } else { "$_" } } | Sort-Object)
                MailboxCount = $(if ($mbxFailed) { $script:LookupFailedText } else { @($mbx | Where-Object { $_.RoleAssignmentPolicy -eq $rapName }).Count })
            }
        })
        Out-Overview -CsvName 'RoleAssignmentPolicies' -Description 'Role assignment policies' -Rows $rapRows -Columns @('Name','IsDefault','MailboxCount') -Failed:$rapFailed
        foreach ($pol in $rapRows) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Description'     = $pol.Description
                'Is default'      = $pol.IsDefault
                'Roles'           = $(if (@($pol.AssignedRoles).Count) { @($pol.AssignedRoles) -join ', ' } else { $null })
                'Mailbox count'   = $pol.MailboxCount
            })
        }

        Add-Line "## Custom management roles"
        Add-Line
        $customRoles = @(); $crFailed = $false
        try { $customRoles = @(Get-ManagementRole | Where-Object { -not (Test-IsTrue $_.IsEndUserRole) -and -not (Test-IsTrue $_.IsRootRole) }) } catch { $crFailed = $true; Add-LogEntry -Section 'Custom management roles' -ErrorRecord $_ }
        $crRows = @($customRoles | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; Parent = $_.Parent; RoleType = $_.RoleType }
        })
        Out-Overview -CsvName 'CustomManagementRoles' -Description 'Custom management roles' -Rows $crRows -Columns @('Name','Parent','RoleType') -Failed:$crFailed

        Add-Line "## Direct user role assignments"
        Add-Line
        $assignments = @(); $asFailed = $false
        try { $assignments = @(Get-ManagementRoleAssignment -RoleAssigneeType User) } catch { $asFailed = $true; Add-LogEntry -Section 'Direct role assignments' -ErrorRecord $_ }
        $asRows = @($assignments | ForEach-Object {
            $recipientWrite = if ($_.CustomRecipientWriteScope) { "Custom scope: $($_.CustomRecipientWriteScope)" }
                              elseif ($_.RecipientAdministrativeUnitScope) { "Administrative unit: $($_.RecipientAdministrativeUnitScope)" }
                              else { "$($_.RecipientWriteScope)" }
            if ("$recipientWrite" -notmatch '\S') { $recipientWrite = $null }
            $configWrite = if ($_.CustomConfigWriteScope) { "Custom scope: $($_.CustomConfigWriteScope)" } else { "$($_.ConfigWriteScope)" }
            [PSCustomObject]@{
                Name = $_.Name; Role = $_.Role; AssignedTo = $_.RoleAssigneeName; AssignmentMethod = $_.AssignmentMethod; Enabled = $_.Enabled
                Scope = $(if ($recipientWrite) { "Recipients: $recipientWrite" } else { 'Not configured' })
                RecipientReadScope = $_.RecipientReadScope; RecipientWriteScope = $recipientWrite
                ConfigReadScope = $_.ConfigReadScope; ConfigWriteScope = $configWrite
                CustomRecipientWriteScope = $_.CustomRecipientWriteScope; CustomConfigWriteScope = $_.CustomConfigWriteScope
                RecipientAdministrativeUnitScope = $_.RecipientAdministrativeUnitScope
            }
        })
        Out-Overview -CsvName 'DirectRoleAssignments' -Description 'Direct user role assignments' -Rows $asRows -Columns @('Name','Role','AssignedTo','Scope') -Failed:$asFailed
        foreach ($a in $asRows) {
            Add-Line "### $(ConvertTo-MdHeading $(if ("$($a.Name)" -match '\S') { $a.Name } else { $a.Role }))"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Role'                   = $a.Role
                'Assigned to'            = $a.AssignedTo
                'Assignment method'      = $a.AssignmentMethod
                'Enabled'                = $a.Enabled
                'Recipient read scope'   = $a.RecipientReadScope
                'Recipient write scope'  = $a.RecipientWriteScope
                'Config read scope'      = $a.ConfigReadScope
                'Config write scope'     = $a.ConfigWriteScope
            })
        }

        Add-Line "## Management scopes"
        Add-Line
        $scopes = @(); $scFailed = $false
        try { $scopes = @(Get-ManagementScope) } catch { $scFailed = $true; Add-LogEntry -Section 'ManagementScope' -ErrorRecord $_ }
        $scRows = @($scopes | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; ScopeRestrictionType = $_.ScopeRestrictionType; Exclusive = $_.Exclusive
                              RecipientRoot = $_.RecipientRoot; RecipientFilter = $_.RecipientFilter }
        })
        Out-Overview -CsvName 'ManagementScopes' -Description 'Management scopes' -Rows $scRows -Columns @('Name','ScopeRestrictionType','Exclusive','RecipientFilter') -Failed:$scFailed

        Add-Line "## RBAC for Applications"
        Add-Line
        $spFailed = $script:SpCacheFailed
        if ($null -eq $script:SpCache -or $spFailed) {
            $script:SpCacheFailed = $false
            try { $script:SpCache = @(Get-ServicePrincipal) } catch { $spFailed = $true; $script:SpCacheFailed = $true; Add-LogEntry -Section 'ServicePrincipal' -ErrorRecord $_ }
        }
        $sps = @($script:SpCache)
        $spAssignments = @(); $spAssignmentsFailed = $false
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
            catch { $spAssignmentsFailed = $true; Add-LogEntry -Section 'SP role assignments' -ErrorRecord $_ }
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
        Register-Csv -Name 'ServicePrincipals' -Description 'Service principals' -Rows $sps -Failed:$spFailed
        Out-Overview -CsvName 'RbacForApplications' -Description 'RBAC for Applications' -Rows $rbacRows -Columns @('DisplayName','Role','AppId','ObjectId') -Failed:$spAssignmentsFailed
        if ($spFailed -and -not $spAssignmentsFailed) { Add-Line "_Service principal names could not be resolved (see collection log); display names may be missing._"; Add-Line }
    }
}

function Write-ComplianceSection {
    Invoke-Section -Title 'Compliance and retention' -Body {
        Add-Line "Messaging records management (MRM) retention, information rights management (IRM) and Office Message Encryption (OME), plus Purview data when collected with -IncludePurview."
        Add-Line

        Add-Line "## MRM retention policies"
        Add-Line
        $rp = @(); $rpFailed = $false
        try { $rp = @(Get-RetentionPolicy) } catch { $rpFailed = $true; Add-LogEntry -Section 'RetentionPolicy' -ErrorRecord $_ }
        $rpRows = @($rp | ForEach-Object {
            [PSCustomObject]@{
                Name = $(if ($_.Name -eq 'ArbitrationMailbox') { 'ArbitrationMailbox (system policy, hidden in the portal)' } else { $_.Name })
                IsDefault = $_.IsDefault
                RetentionPolicyTagLinks = @($_.RetentionPolicyTagLinks | ForEach-Object { if ($_ -is [string]) { $_ } elseif ($_.Name) { "$($_.Name)" } else { "$_" } } | Sort-Object)
            }
        })
        Out-Overview -CsvName 'RetentionPolicies' -Description 'MRM retention policies' -Rows $rpRows -Columns @('Name','IsDefault') -Failed:$rpFailed
        foreach ($pol in $rpRows) {
            Add-Line "### $(ConvertTo-MdHeading $pol.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Is default'     = $pol.IsDefault
                'Retention tags' = $(if (@($pol.RetentionPolicyTagLinks).Count) { @($pol.RetentionPolicyTagLinks) -join ', ' } else { $null })
            })
        }
        Add-Line "## MRM retention tags"
        Add-Line
        $tags = @(); $tagsFailed = $false
        try { $tags = @(Get-RetentionPolicyTag) } catch { $tagsFailed = $true; Add-LogEntry -Section 'RetentionPolicyTag' -ErrorRecord $_ }
        $tagActionMap = @{ DeleteAndAllowRecovery = 'Delete and allow recovery'; PermanentlyDelete = 'Permanently delete'; MoveToArchive = 'Move to archive'; MarkAsPastRetentionLimit = 'Mark as past retention limit' }
        $tagTypeMap = @{ All = 'Default (entire mailbox)'; RecoverableItems = 'Recoverable Items'; DeletedItems = 'Deleted Items'; JunkEmail = 'Junk Email' }
        $tagRows = @($tags | Where-Object { -not $_.SystemTag } | ForEach-Object {
            $tagName = $_.Name
            [PSCustomObject]@{
                Name = $tagName
                Type = $(if ($tagTypeMap.ContainsKey("$($_.Type)")) { $tagTypeMap["$($_.Type)"] } else { $_.Type })
                RetentionPeriod = Format-RetentionPeriod -Tag $_
                RetentionAction = $(if ($tagActionMap.ContainsKey("$($_.RetentionAction)")) { $tagActionMap["$($_.RetentionAction)"] } else { $_.RetentionAction })
                RawType = "$($_.Type)"; RawAction = "$($_.RetentionAction)"
                UsedByPolicies = @($rpRows | Where-Object { @($_.RetentionPolicyTagLinks) -contains $tagName } | ForEach-Object { $_.Name })
                Comment = $_.Comment
            }
        } | Sort-Object Type, Name)
        Out-Overview -CsvName 'RetentionPolicyTags' -Description 'MRM retention tags (system tags excluded)' -Rows $tagRows -Columns @('Name','Type','RetentionPeriod','RetentionAction') -Failed:$tagsFailed

        Add-Line "## IRM and OME configuration"
        Add-Line
        $irm = @(); $irmFailed = $false
        try { $irm = @(Get-IRMConfiguration) } catch { $irmFailed = $true; Add-LogEntry -Section 'IRMConfiguration' -ErrorRecord $_ }
        Register-Csv -Name 'IrmConfiguration' -Description 'IRM configuration' -Rows $irm -Failed:$irmFailed
        if ($irm.Count -gt 0) {
            Out-Card -Settings ([ordered]@{
                'Internal licensing'              = $irm[0].InternalLicensingEnabled
                'External licensing'              = $irm[0].ExternalLicensingEnabled
                'Azure RMS licensing'             = $irm[0].AzureRMSLicensingEnabled
                'Transport decryption'            = $irm[0].TransportDecryptionSetting
                'Journal report decryption'       = $irm[0].JournalReportDecryptionEnabled
                'Search enabled'                  = $irm[0].SearchEnabled
            })
        }
        elseif ($irmFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not configured._"; Add-Line }
        $ome = @(); $omeFailed = $false
        try { $ome = @(Get-OMEConfiguration) } catch { $omeFailed = $true; Add-LogEntry -Section 'OMEConfiguration' -ErrorRecord $_ }
        Register-Csv -Name 'OmeConfiguration' -Description 'OME configuration' -Rows $ome -Failed:$omeFailed
        if ($ome.Count -gt 0) {
            Out-Card -Settings ([ordered]@{
                'OTP enabled'               = $ome[0].OTPEnabled
                'Social ID sign-in'         = $ome[0].SocialIdSignIn
                'External mail expiry (days)' = $ome[0].ExternalMailExpiryInDays
            })
        }
        elseif ($omeFailed) { Add-Line "_Not available - lookup failed (see collection log)._"; Add-Line }
        else { Add-Line "_Not configured._"; Add-Line }
    }

    Invoke-Section -Title 'Purview retention policies covering Exchange' -Level 2 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Purview retention') { return }
        Assert-Cmdlet Get-RetentionCompliancePolicy
        $policies = @(Get-RetentionCompliancePolicy -DistributionDetail -WarningAction SilentlyContinue)
        $tagNames = @{}
        try {
            foreach ($t in @(Get-ComplianceTag -WarningAction SilentlyContinue)) {
                foreach ($k in @("$($t.Guid)", "$($t.ImmutableId)", "$($t.Identity)", "$($t.Name)")) {
                    if ($k -and -not $tagNames.ContainsKey($k)) { $tagNames[$k] = "$($t.Name)" }
                }
            }
        }
        catch { Add-LogEntry -Section 'ComplianceTag (name map)' -ErrorRecord $_ }
        $exo = @($policies | Where-Object { (Get-NonEmptyCount $_.ExchangeLocation) -gt 0 })
        $rows = @($exo | ForEach-Object {
            $pn = $_.Name
            $rules = @()
            try { $rules = @(@(Get-RetentionComplianceRule -Policy $pn -WarningAction SilentlyContinue) | ForEach-Object { Format-RetentionRule -Rule $_ -TagNames $tagNames } | Sort-Object -Unique) }
            catch { Add-LogEntry -Section "RetentionComplianceRule ($pn)" -ErrorRecord $_; $rules = @('Unknown (rule lookup failed)') }
            [PSCustomObject]@{
                Name = $pn; Status = $(if (Test-IsTrue $_.Enabled) { 'On' } else { 'Off' }); Mode = $_.Mode; Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation; ExchangeLocationException = $_.ExchangeLocationException
                Rules = $rules
            }
        })
        Out-Overview -CsvName 'PurviewRetentionPolicies' -Description 'Purview retention policies covering Exchange' -Rows $rows -Columns @('Name','Status','Mode')
        foreach ($p in $rows) {
            Add-Line "### $(ConvertTo-MdHeading $p.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Status'                        = $p.Status
                'Mode'                          = $p.Mode
                'Exchange locations'            = @{ Value = $p.ExchangeLocation; CsvFile = 'PurviewRetentionPolicies.csv' }
                'Exchange location exceptions'  = @{ Value = $p.ExchangeLocationException; CsvFile = 'PurviewRetentionPolicies.csv' }
                'Retention settings'            = $(if (@($p.Rules).Count) { @($p.Rules) -join '; ' } else { $null })
            })
        }
    }

    Invoke-Section -Title 'DLP policies covering Exchange' -Level 2 -Body {
        if (Add-PurviewStatusOrThrow -Section 'DLP policies') { return }
        Assert-Cmdlet Get-DlpCompliancePolicy
        $policies = @(Get-DlpCompliancePolicy -WarningAction SilentlyContinue)
        $exo = @($policies | Where-Object { (Get-NonEmptyCount $_.ExchangeLocation) -gt 0 })
        $modeMap = @{ Enable = 'On'; TestWithNotifications = 'Test with notifications (simulation)'; TestWithoutNotifications = 'Test without notifications (simulation)'; Disable = 'Off' }
        $rows = @($exo | ForEach-Object {
            $pn = $_.Name
            $rules = @(); $rulesFailed = $false
            try { $rules = @(Get-DlpComplianceRule -Policy $pn -WarningAction SilentlyContinue) } catch { $rulesFailed = $true; Add-LogEntry -Section "DlpComplianceRule ($pn)" -ErrorRecord $_ }
            $ruleLines = @('Unknown (rule lookup failed)')
            if (-not $rulesFailed) {
                $ruleLines = @($rules | ForEach-Object {
                    $acts = @()
                    if (Test-IsTrue $_.BlockAccess) { $acts += 'blocks access' }
                    if ((Get-NonEmptyCount $_.NotifyUser) -gt 0) { $acts += 'notifies user' }
                    if ((Get-NonEmptyCount $_.GenerateAlert) -gt 0) { $acts += 'generates alert' }
                    if ((Get-NonEmptyCount $_.GenerateIncidentReport) -gt 0) { $acts += 'incident report' }
                    $actText = if ($acts.Count) { $acts -join ', ' } else { 'no actions' }
                    "$($_.Name): $actText$(if (Test-IsTrue $_.Disabled) { ' (rule off)' } else { '' })"
                })
            }
            $modeText = if ($modeMap.ContainsKey("$($_.Mode)")) { $modeMap["$($_.Mode)"] } else { "$($_.Mode)" }
            [PSCustomObject]@{
                Name = $pn; Mode = $modeText; Enabled = $_.Enabled
                ExchangeLocation = $_.ExchangeLocation; ExchangeLocationException = $_.ExchangeLocationException
                Rules = $ruleLines
            }
        })
        Out-Overview -CsvName 'DlpPolicies' -Description 'DLP policies covering Exchange' -Rows $rows -Columns @('Name','Mode','ExchangeLocation')
        foreach ($p in $rows) {
            Add-Line "### $(ConvertTo-MdHeading $p.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Mode'                        = $p.Mode
                'Exchange locations'          = @{ Value = $p.ExchangeLocation; CsvFile = 'DlpPolicies.csv' }
                'Exchange location exceptions' = @{ Value = $p.ExchangeLocationException; CsvFile = 'DlpPolicies.csv' }
                'Rules'                       = @{ Value = $p.Rules; CsvFile = 'DlpPolicies.csv' }
            })
        }
    }

    Invoke-Section -Title 'Sensitivity labels and label policies' -Level 2 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Sensitivity labels') { return }
        Assert-Cmdlet Get-Label
        $labels = @(Get-Label -WarningAction SilentlyContinue)
        $labelRows = @($labels | Sort-Object Priority | ForEach-Object {
            $li = Get-SensitivityLabelInfo -Label $_ -AllLabels $labels
            [PSCustomObject]@{
                DisplayName = $_.DisplayName; Name = $_.Name; Priority = $_.Priority; Scope = $li.Scope; ParentLabel = $li.Parent
                Sublabels = $li.Sublabels; DescriptionForUsers = $_.Tooltip; Description = $_.Comment
                AccessControl = $li.AccessControl; ContentMarking = $li.ContentMarking; AutoLabeling = $li.AutoLabeling
                GroupSettings = $li.GroupSettings; SiteSettings = $li.SiteSettings
                LabelActionTypes = $li.ActionTypes; ContentType = $_.ContentType; Disabled = $_.Disabled
            }
        })
        Out-Overview -CsvName 'SensitivityLabels' -Description 'Sensitivity labels' -Rows $labelRows -Columns @('DisplayName','Priority','Scope','ParentLabel')
        foreach ($l in $labelRows) {
            Add-Line "### $(ConvertTo-MdHeading $l.DisplayName)"
            Add-Line
            $card = [ordered]@{
                'Name'                                = $l.Name
                'Display name'                        = $l.DisplayName
                'Priority'                            = $l.Priority
                'Parent label'                        = $l.ParentLabel
            }
            if (@($l.Sublabels).Count) { $card['Sublabels'] = @($l.Sublabels) -join ', ' }
            $card['Description for users']                = $l.DescriptionForUsers
            $card['Description']                          = $l.Description
            $card['Scope']                                = $l.Scope
            $card['Access control']                       = $l.AccessControl
            $card['Content marking']                      = $l.ContentMarking
            $card['Auto-labeling for files and emails']   = $l.AutoLabeling
            $card['Group settings']                       = $l.GroupSettings
            $card['Site settings']                        = $l.SiteSettings
            Out-Card -Settings $card
        }
        $labelPolicies = @(); $lpFailed = $false
        try { $labelPolicies = @(Get-LabelPolicy -WarningAction SilentlyContinue) } catch { $lpFailed = $true; Add-LogEntry -Section 'LabelPolicy' -ErrorRecord $_ }
        $lpRows = @($labelPolicies | ForEach-Object {
            $pub = @(@($_.ExchangeLocation) + @($_.ModernGroupLocation) | Where-Object { "$_" -match '\S' })
            $exc = @(@($_.ExchangeLocationException) + @($_.ModernGroupLocationException) | Where-Object { "$_" -match '\S' })
            [PSCustomObject]@{
                Name = $_.Name; Description = $_.Comment
                PublishedLabels = @(Resolve-LabelNames -Ids @($_.Labels) -AllLabels $labels)
                PublishedTo = $pub; PublishedToExceptions = $exc
                ExchangeLocation = $_.ExchangeLocation; ExchangeLocationException = $_.ExchangeLocationException
                ModernGroupLocation = $_.ModernGroupLocation; ModernGroupLocationException = $_.ModernGroupLocationException
                Enabled = $_.Enabled
            }
        })
        Add-Line "### Label policies"
        Add-Line
        Out-Overview -CsvName 'LabelPolicies' -Description 'Sensitivity label policies' -Rows @($lpRows | ForEach-Object {
            [PSCustomObject]@{ Name = $_.Name; PublishedLabels = @($_.PublishedLabels).Count; PublishedTo = (@($_.PublishedTo) -join ', ') }
        }) -Columns @('Name','PublishedLabels','PublishedTo') -SkipCsv -Failed:$lpFailed
        Register-Csv -Name 'LabelPolicies' -Description 'Sensitivity label policies' -Rows $lpRows -Failed:$lpFailed
        foreach ($p in $lpRows) {
            Add-Line "#### $(ConvertTo-MdHeading $p.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Description'                  = $p.Description
                'Published labels'             = $(if (@($p.PublishedLabels).Count) { @($p.PublishedLabels) -join ', ' } else { $null })
                'Published to users and groups' = @{ Value = $p.PublishedTo; CsvFile = 'LabelPolicies.csv' }
                'Excluded users and groups'    = @{ Value = $p.PublishedToExceptions; CsvFile = 'LabelPolicies.csv' }
            })
        }
    }

    Invoke-Section -Title 'Retention labels' -Level 2 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Retention labels') { return }
        Assert-Cmdlet Get-ComplianceTag
        $tags = @(Get-ComplianceTag -WarningAction SilentlyContinue)
        $actMap = @{ Keep = 'Retain'; Delete = 'Delete'; KeepAndDelete = 'Retain, then delete' }
        $typeMap = @{ CreationAgeInDays = 'When created'; ModificationAgeInDays = 'When last modified'; EventAgeInDays = 'Based on event'; TaggedAgeInDays = 'When labelled' }
        $rows = @($tags | ForEach-Object {
            $dur = "$($_.RetentionDuration)"
            $d = 0
            $durText = if ($dur -eq '' -or $dur -eq 'Unlimited') { 'Forever' }
                       elseif ([int]::TryParse($dur, [ref]$d)) { if ($d -gt 0 -and $d % 365 -eq 0) { "$d days ($($d / 365) year(s))" } else { "$d days" } }
                       else { $dur }
            [PSCustomObject]@{ Name = $_.Name; RetentionPeriod = $durText
                              RetentionAction = $(if ($actMap.ContainsKey("$($_.RetentionAction)")) { $actMap["$($_.RetentionAction)"] } else { "$($_.RetentionAction)" })
                              RetentionType = $(if ($typeMap.ContainsKey("$($_.RetentionType)")) { $typeMap["$($_.RetentionType)"] } else { "$($_.RetentionType)" })
                              IsRecordLabel = $_.IsRecordLabel; EventType = $_.EventType; Notes = $_.Notes
                              RawRetentionDuration = $_.RetentionDuration; RawRetentionAction = "$($_.RetentionAction)" }
        })
        Out-Overview -CsvName 'ComplianceTags' -Description 'Retention labels (compliance tags)' -Rows $rows -Columns @('Name','RetentionPeriod','RetentionAction','IsRecordLabel')
        foreach ($t in $rows) {
            Add-Line "### $(ConvertTo-MdHeading $t.Name)"
            Add-Line
            Out-Card -Settings ([ordered]@{
                'Retention period' = $t.RetentionPeriod
                'Action'           = $t.RetentionAction
                'Based on'         = $t.RetentionType
                'Record label'     = $t.IsRecordLabel
                'Event type'       = $t.EventType
                'Notes'            = $t.Notes
            })
        }
    }

    Invoke-Section -Title 'Alert policies' -Level 2 -Body {
        if (Add-PurviewStatusOrThrow -Section 'Alert policies') { return }
        Assert-Cmdlet Get-ProtectionAlert
        $alerts = @(Get-ProtectionAlert -WarningAction SilentlyContinue)
        $all = @($alerts | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name; Type = $(if (Test-IsTrue $_.IsSystemRule) { 'System' } else { 'Custom' }); Category = $_.Category; Severity = $_.Severity
                Status = $(if (Test-IsTrue $_.Disabled) { 'Off' } else { 'On' })
                ManagedIn = $(if ("$($_.Category)" -eq 'DataLossPrevention') { 'Purview DLP policy' } else { 'Defender portal (Alert policy)' })
                NotifyUser = $_.NotifyUser; ThreatType = $_.ThreatType; Disabled = $_.Disabled
            }
        })
        Register-Csv -Name 'ProtectionAlerts' -Description 'All alert policies (system and custom)' -Rows $all
        $dlp = @($all | Where-Object { $_.ManagedIn -eq 'Purview DLP policy' })
        $portal = @($all | Where-Object { $_.ManagedIn -ne 'Purview DLP policy' })
        Out-Card -Settings ([ordered]@{
            'Alert policies (total, PowerShell)'              = $all.Count
            'Custom alert policies'                           = @($portal | Where-Object { $_.Type -eq 'Custom' }).Count
            'System alert policies'                           = @($portal | Where-Object { $_.Type -eq 'System' }).Count
            'DLP alert policies (created by DLP policy rules)' = $dlp.Count
        })
        Add-Line "The Defender portal (Email & collaboration > Policies & rules > Alert policy) does not list DLP alert policies, which are managed through their DLP policy in Microsoft Purview, and hides a few system policies. Its count can therefore be lower than the PowerShell total. The full list is in ``ProtectionAlerts.csv``."
        Add-Line
        Add-Line "### Custom alert policies"
        Add-Line
        Out-Overview -CsvName 'ProtectionAlertsCustom' -Description 'Custom alert policies' -Rows @($portal | Where-Object { $_.Type -eq 'Custom' }) -Columns @('Name','Category','Severity','Status')
        Add-Line "### DLP alert policies"
        Add-Line
        Out-Overview -CsvName 'ProtectionAlertsDlp' -Description 'DLP alert policies' -Rows $dlp -Columns @('Name','Severity','Status')
    }
}

function Write-AppendixSection {
    Invoke-Section -Title 'Appendix' -Body {
        Add-Line "Supplementary information: the CSV file index and the collection log."
        Add-Line

        Add-Line "## A. CSV file index"
        Add-Line
        $indexRows = @($script:CsvIndex.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{ File = $_.Value.File; Contents = $_.Value.Contents; Rows = $_.Value.Rows }
        })
        if ($indexRows.Count -eq 0) { Add-Line "_No CSV files._"; Add-Line }
        else {
            Add-Line "| File | Contents | Rows |"
            Add-Line "|---|---|---|"
            foreach ($r in $indexRows) { Add-Line "| ``$($r.File)`` | $(Format-Value $r.Contents) | $($r.Rows) |" }
            Add-Line
        }

        Add-Line "## B. Collection log"
        Add-Line
        if ($script:CollectionLog.Count -eq 0) {
            Add-Line "No skipped or failed sections."
            Add-Line
        }
        else {
            Add-Line "| Section | Reason | Time |"
            Add-Line "|---|---|---|"
            foreach ($entry in $script:CollectionLog) {
                Add-Line "| $(Format-Value $entry.Section) | $(Format-Value $entry.Reason) | $($entry.Time) |"
            }
            Add-Line
        }

        Add-Line "## Parameters used"
        Add-Line
        Out-Card -Settings ([ordered]@{
            'CustomerName'        = $CustomerName
            'Author'              = $Author
            'DocumentDate'        = $DocumentDate.ToString('yyyy-MM-dd')
            'OutputPath'          = $OutputPath
            'UserPrincipalName'   = $UserPrincipalName
            'AppId'               = $(if ($AppId) { $AppId } else { '(interactive)' })
            'IncludePurview'      = [bool]$IncludePurview
            'IncludeIntegratedApps' = [bool]$IncludeIntegratedApps
            'MaxListItems'        = $MaxListItems
            'MaxTableRows'        = $MaxTableRows
            'CsvDelimiter'        = $(if ($CsvDelimiter) { $CsvDelimiter } else { '(UseCulture)' })
            'DnsServer'           = $(if ($DnsServer) { $DnsServer } else { '(default)' })
            'ExoLogPath'          = $ExoLogPath
        })

        $duration = (Get-Date) - $script:RunStart
        Add-Line "## Run duration"
        Add-Line
        Add-Line "$([math]::Round($duration.TotalMinutes, 1)) minutes ($($duration.ToString('hh\:mm\:ss')))"
        Add-Line
    }
}

# ============================================================
# Main
# ============================================================

$safeCustomer = ($CustomerName -replace '[\\/:*?"<>|\[\]\s]', '_')
$dateStamp = $DocumentDate.ToString('yyyyMMdd')
$docName = "EXO-Documentation_${safeCustomer}_${dateStamp}.md"
$resolvedOutput = Resolve-Path -LiteralPath $OutputPath -ErrorAction SilentlyContinue
if (-not $resolvedOutput) {
    New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
    $resolvedOutput = Resolve-Path -LiteralPath $OutputPath
}
$docPath = Join-Path $resolvedOutput.Path $docName
$csvDir = Join-Path $resolvedOutput.Path "EXO-Documentation_${safeCustomer}_${dateStamp}_csv"

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
        Register-ExoProxyWrapper -ExcludeConnectionIds $script:PreExistingConnectionIds
        if ($script:PreExistingConnectionIds.Count -gt 0) { Write-Warning "Other Exchange Online sessions were already open; only the session opened by this script is used for collection." }
        Write-Host "Connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        $messages = @(); $ex = $_.Exception
        while ($ex) { $messages += $ex.Message; $ex = $ex.InnerException }
        Write-Error "Failed to connect to Exchange Online: $(($messages | Select-Object -Unique) -join ' --> ')"
        exit 1
    }

    # Pre-collect tenant identity for the document information section
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
        $ours = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $script:PreExistingConnectionIds -notcontains $_.ConnectionId })
        $connInfo = @(@($ours | Where-Object { -not (Test-IsTrue $_.IsEopSession) }) + @($ours) | Select-Object -First 1)[0]
        if ($connInfo.UserPrincipalName) { $script:CollectedBy = $connInfo.UserPrincipalName }
        elseif ($connInfo.User) { $script:CollectedBy = $connInfo.User }
        if ($connInfo.TenantID) { $script:TenantId = $connInfo.TenantID }
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
            Register-ExoProxyWrapper -ExcludeConnectionIds $script:PreExistingConnectionIds
            Write-Host "Connected to Security & Compliance." -ForegroundColor Green
        }
        catch {
            $script:PurviewError = Get-ErrorDetail $_
            Write-Warning "Could not connect to Security & Compliance PowerShell - $script:PurviewError. Purview sections will be marked Not available."
            Add-LogEntry -Section 'Purview connection' -Reason $script:PurviewError
        }
    }

    if ($IncludeIntegratedApps) {
        if ($AppId) {
            $script:IntegratedAppsError = 'Not supported with app-only authentication (-AppId); run interactively'
        }
        elseif (-not (Get-Module -ListAvailable O365CentralizedAddInDeployment)) {
            $script:IntegratedAppsError = 'O365CentralizedAddInDeployment module not installed (Install-Module O365CentralizedAddInDeployment -Scope CurrentUser)'
        }
        else {
            Write-Host "Connecting to the Microsoft 365 add-in service (Integrated apps)..." -ForegroundColor Cyan
            try {
                Import-Module O365CentralizedAddInDeployment -ErrorAction Stop
                Connect-OrganizationAddInService -ErrorAction Stop
                $script:IntegratedAppsConnected = $true
                Write-Host "Connected to the add-in service." -ForegroundColor Green
                # The module keeps its session in a private variable set in the caller's
                # scope, so collection must happen in this same (script) scope.
                try { $script:IntegratedAppsList = @(Get-OrganizationAddIn -ErrorAction Stop) }
                catch { $script:IntegratedAppsListError = $_ }
                foreach ($ia in @($script:IntegratedAppsList)) {
                    if (-not $ia.ProductId) { continue }
                    try { $script:IntegratedAppsDetails["$($ia.ProductId)"] = Get-OrganizationAddIn -ProductId $ia.ProductId -ErrorAction Stop }
                    catch { $script:IntegratedAppsDetailErrors["$($ia.ProductId)"] = $_ }
                }
            }
            catch { $script:IntegratedAppsError = Get-ErrorDetail $_ }
        }
        if (-not $script:IntegratedAppsConnected) {
            Write-Warning "Could not connect to the Microsoft 365 add-in service - $script:IntegratedAppsError. The Integrated apps section will be marked Not available."
            Add-LogEntry -Section 'Integrated apps connection' -Reason $script:IntegratedAppsError
        }
    }

    Write-Host "`nCollecting documentation data..." -ForegroundColor Cyan
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

    # Write Markdown document
    $script:Report.ToString() | Out-File -LiteralPath $docPath -Encoding utf8
    Write-Host "`nDocument written to: $docPath" -ForegroundColor Green

    # CSV export (always on, all rows and all collected properties)
    if ($script:CsvData.Count -gt 0) {
        New-Item -ItemType Directory -Force -Path $csvDir | Out-Null
        foreach ($key in $script:CsvData.Keys) {
            $csvPath = Join-Path $csvDir "$key.csv"
            if ($CsvDelimiter) {
                @(ConvertTo-CsvRow -Rows $script:CsvData[$key]) | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -Delimiter $CsvDelimiter
            }
            else {
                @(ConvertTo-CsvRow -Rows $script:CsvData[$key]) | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -UseCulture
            }
        }
        Write-Host "CSV exports written to: $csvDir" -ForegroundColor Green
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Host "`nDisconnecting sessions opened by this script..." -ForegroundColor Cyan
    try {
        $canTarget = $false
        $discCmd = Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue
        if ($discCmd) { $canTarget = $discCmd.Parameters.ContainsKey('ConnectionId') }
        if ($canTarget) {
            $current = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.ConnectionId })
            $ours = @($current | Where-Object { $script:PreExistingConnectionIds -notcontains $_.ConnectionId })
            if ($ours.Count -gt 0) {
                foreach ($c in $ours) {
                    Disconnect-ExchangeOnline -ConnectionId $c.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue
                }
            }
            elseif ($current.Count -eq 0 -and @($script:PreExistingConnectionIds).Count -eq 0) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        elseif (@($script:PreExistingConnectionIds).Count -eq 0) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
        else {
            Write-Warning "The installed ExchangeOnlineManagement module cannot disconnect individual sessions; sessions opened by this script were left connected to preserve pre-existing sessions."
        }
    }
    catch { Write-Warning "Disconnect failed: $_" }
    Write-Host "Done." -ForegroundColor Green
}
