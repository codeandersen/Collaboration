<#
.SYNOPSIS
    Gets all transport rules (mail flow rules) and their configuration from Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all transport rules
    with their complete configuration settings including conditions, actions,
    and exceptions, then exports them to a CSV file.

.PARAMETER ExportPath
    Path to export the results to a CSV file. If not specified, defaults to
    "TransportRules_<timestamp>.csv" in the current directory.

.EXAMPLE
    .\Get-TransportRules.ps1
    Exports transport rules to a timestamped CSV file in the current directory.

.EXAMPLE
    .\Get-TransportRules.ps1 -ExportPath "C:\Reports\TransportRules.csv"
    Exports transport rules to a specific CSV file.

.NOTES
    Requires: ExchangeOnlineManagement module
    Install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "TransportRules_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

# Function to check and install ExchangeOnlineManagement module
function Initialize-ExchangeOnlineModule {
    Write-Host "Checking for ExchangeOnlineManagement module..." -ForegroundColor Cyan
    
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
            Write-Host "Module installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install ExchangeOnlineManagement module: $_"
            exit 1
        }
    }
    else {
        Write-Host "ExchangeOnlineManagement module found." -ForegroundColor Green
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineService {
    Write-Host "`nConnecting to Exchange Online..." -ForegroundColor Cyan
    
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

# Function to convert array/list properties to string
function Convert-ArrayToString {
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [array] -or $Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return ($Value -join '; ')
    }
    return $Value.ToString()
}

# Main script execution
try {
    # Ensure module is installed
    Initialize-ExchangeOnlineModule
    
    # Connect to Exchange Online
    Connect-ExchangeOnlineService
    
    # Get all transport rules
    Write-Host "`nRetrieving all transport rules..." -ForegroundColor Cyan
    $transportRules = Get-TransportRule
    
    if ($transportRules.Count -eq 0) {
        Write-Host "No transport rules found." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($transportRules.Count) transport rule(s)." -ForegroundColor Green
    
    # Array to store results
    $results = @()
    
    # Process each transport rule
    Write-Host "`nProcessing transport rules..." -ForegroundColor Cyan
    
    foreach ($rule in $transportRules) {
        Write-Host "  Processing: $($rule.Name)..." -ForegroundColor Gray
        
        $results += [PSCustomObject]@{
            Name                                = $rule.Name
            State                               = $rule.State
            Priority                            = $rule.Priority
            Mode                                = $rule.Mode
            Comments                            = $rule.Comments
            Description                         = $rule.Description
            
            # Conditions
            From                                = Convert-ArrayToString $rule.From
            FromMemberOf                        = Convert-ArrayToString $rule.FromMemberOf
            FromScope                           = $rule.FromScope
            SentTo                              = Convert-ArrayToString $rule.SentTo
            SentToMemberOf                      = Convert-ArrayToString $rule.SentToMemberOf
            SentToScope                         = $rule.SentToScope
            SubjectContainsWords                = Convert-ArrayToString $rule.SubjectContainsWords
            SubjectOrBodyContainsWords          = Convert-ArrayToString $rule.SubjectOrBodyContainsWords
            SubjectMatchesPatterns              = Convert-ArrayToString $rule.SubjectMatchesPatterns
            SubjectOrBodyMatchesPatterns        = Convert-ArrayToString $rule.SubjectOrBodyMatchesPatterns
            HeaderContainsWords                 = Convert-ArrayToString $rule.HeaderContainsWords
            HeaderMatchesPatterns               = Convert-ArrayToString $rule.HeaderMatchesPatterns
            FromAddressContainsWords            = Convert-ArrayToString $rule.FromAddressContainsWords
            FromAddressMatchesPatterns          = Convert-ArrayToString $rule.FromAddressMatchesPatterns
            RecipientAddressContainsWords       = Convert-ArrayToString $rule.RecipientAddressContainsWords
            RecipientAddressMatchesPatterns     = Convert-ArrayToString $rule.RecipientAddressMatchesPatterns
            AttachmentNameMatchesPatterns       = Convert-ArrayToString $rule.AttachmentNameMatchesPatterns
            AttachmentExtensionMatchesWords     = Convert-ArrayToString $rule.AttachmentExtensionMatchesWords
            AttachmentSizeOver                  = $rule.AttachmentSizeOver
            MessageSizeOver                     = $rule.MessageSizeOver
            SCLOver                             = $rule.SCLOver
            WithImportance                      = $rule.WithImportance
            MessageTypeMatches                  = $rule.MessageTypeMatches
            RecipientDomainIs                   = Convert-ArrayToString $rule.RecipientDomainIs
            SenderDomainIs                      = Convert-ArrayToString $rule.SenderDomainIs
            
            # Actions
            PrependSubject                      = $rule.PrependSubject
            SetSCL                              = $rule.SetSCL
            SetHeaderName                       = $rule.SetHeaderName
            SetHeaderValue                      = $rule.SetHeaderValue
            RemoveHeader                        = $rule.RemoveHeader
            ApplyClassification                 = $rule.ApplyClassification
            ApplyHtmlDisclaimerLocation         = $rule.ApplyHtmlDisclaimerLocation
            ApplyHtmlDisclaimerText             = $rule.ApplyHtmlDisclaimerText
            ApplyHtmlDisclaimerFallbackAction   = $rule.ApplyHtmlDisclaimerFallbackAction
            ApplyRightsProtectionTemplate       = $rule.ApplyRightsProtectionTemplate
            ApplyOME                            = $rule.ApplyOME
            RemoveOME                           = $rule.RemoveOME
            RedirectMessageTo                   = Convert-ArrayToString $rule.RedirectMessageTo
            RejectMessageEnhancedStatusCode     = $rule.RejectMessageEnhancedStatusCode
            RejectMessageReasonText             = $rule.RejectMessageReasonText
            DeleteMessage                       = $rule.DeleteMessage
            Quarantine                          = $rule.Quarantine
            BlindCopyTo                         = Convert-ArrayToString $rule.BlindCopyTo
            ModerateMessageByUser               = Convert-ArrayToString $rule.ModerateMessageByUser
            ModerateMessageByManager            = $rule.ModerateMessageByManager
            RouteMessageOutboundConnector       = $rule.RouteMessageOutboundConnector
            RouteMessageOutboundRequireTls      = $rule.RouteMessageOutboundRequireTls
            StopRuleProcessing                  = $rule.StopRuleProcessing
            SenderAddressLocation               = $rule.SenderAddressLocation
            
            # Exceptions
            ExceptIfFrom                        = Convert-ArrayToString $rule.ExceptIfFrom
            ExceptIfFromMemberOf                = Convert-ArrayToString $rule.ExceptIfFromMemberOf
            ExceptIfSentTo                      = Convert-ArrayToString $rule.ExceptIfSentTo
            ExceptIfSentToMemberOf              = Convert-ArrayToString $rule.ExceptIfSentToMemberOf
            ExceptIfSubjectContainsWords        = Convert-ArrayToString $rule.ExceptIfSubjectContainsWords
            ExceptIfRecipientDomainIs           = Convert-ArrayToString $rule.ExceptIfRecipientDomainIs
            ExceptIfSenderDomainIs              = Convert-ArrayToString $rule.ExceptIfSenderDomainIs
            
            # Metadata
            Identity                            = $rule.Identity
            Guid                                = $rule.Guid
            WhenChanged                         = $rule.WhenChanged
            ActivationDate                      = $rule.ActivationDate
            ExpiryDate                          = $rule.ExpiryDate
            DistinguishedName                   = $rule.DistinguishedName
        }
    }
    
    # Display results summary
    Write-Host "`n==== TRANSPORT RULES SUMMARY ====" -ForegroundColor Cyan
    $results | Select-Object Name, State, Priority, Mode, Description | Format-Table -AutoSize
    
    # Export to CSV
    Write-Host "`nExporting results to: $ExportPath" -ForegroundColor Cyan
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export completed successfully." -ForegroundColor Green
    
    # Summary statistics
    Write-Host "`n==== STATISTICS ====" -ForegroundColor Cyan
    Write-Host "Total Transport Rules: $($results.Count)" -ForegroundColor Green
    Write-Host "Enabled Rules: $(($results | Where-Object {$_.State -eq 'Enabled'}).Count)" -ForegroundColor Green
    Write-Host "Disabled Rules: $(($results | Where-Object {$_.State -eq 'Disabled'}).Count)" -ForegroundColor Yellow
    
    Write-Host "`nRule Modes:" -ForegroundColor Yellow
    $results | Group-Object -Property Mode | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor Gray
    }
    
    Write-Host "`nTop 5 Rules by Priority:" -ForegroundColor Yellow
    $results | Sort-Object Priority | Select-Object -First 5 | ForEach-Object {
        Write-Host "  Priority $($_.Priority): $($_.Name) [$($_.State)]" -ForegroundColor Gray
    }
    
    # Rules with expiry dates
    $rulesWithExpiry = $results | Where-Object {$null -ne $_.ExpiryDate}
    if ($rulesWithExpiry.Count -gt 0) {
        Write-Host "`nRules with Expiry Dates: $($rulesWithExpiry.Count)" -ForegroundColor Yellow
        $rulesWithExpiry | ForEach-Object {
            Write-Host "  $($_.Name): Expires on $($_.ExpiryDate)" -ForegroundColor Gray
        }
    }
    
    Write-Host "`nExport Location: $ExportPath" -ForegroundColor Cyan
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    # Disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Disconnected." -ForegroundColor Green
}
