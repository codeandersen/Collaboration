<#
.SYNOPSIS
    Gets all outbound connectors and their configuration from Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all outbound connectors
    with their complete configuration settings and exports them to a CSV file.

.PARAMETER ExportPath
    Path to export the results to a CSV file. If not specified, defaults to
    "OutboundConnectors_<timestamp>.csv" in the current directory.

.EXAMPLE
    .\Get-OutboundConnectors.ps1
    Exports outbound connectors to a timestamped CSV file in the current directory.

.EXAMPLE
    .\Get-OutboundConnectors.ps1 -ExportPath "C:\Reports\OutboundConnectors.csv"
    Exports outbound connectors to a specific CSV file.

.NOTES
    Requires: ExchangeOnlineManagement module
    Install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "OutboundConnectors_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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

# Main script execution
try {
    # Ensure module is installed
    Initialize-ExchangeOnlineModule
    
    # Connect to Exchange Online
    Connect-ExchangeOnlineService
    
    # Get all outbound connectors
    Write-Host "`nRetrieving all outbound connectors..." -ForegroundColor Cyan
    $outboundConnectors = Get-OutboundConnector
    
    if ($outboundConnectors.Count -eq 0) {
        Write-Host "No outbound connectors found." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($outboundConnectors.Count) outbound connector(s)." -ForegroundColor Green
    
    # Array to store results
    $results = @()
    
    # Process each outbound connector
    Write-Host "`nProcessing outbound connectors..." -ForegroundColor Cyan
    
    foreach ($connector in $outboundConnectors) {
        Write-Host "  Processing: $($connector.Name)..." -ForegroundColor Gray
        
        $results += [PSCustomObject]@{
            Name                          = $connector.Name
            Enabled                       = $connector.Enabled
            ConnectorType                 = $connector.ConnectorType
            ConnectorSource               = $connector.ConnectorSource
            Comment                       = $connector.Comment
            RecipientDomains              = ($connector.RecipientDomains -join '; ')
            SmartHosts                    = ($connector.SmartHosts -join '; ')
            TlsSettings                   = $connector.TlsSettings
            TlsDomain                     = $connector.TlsDomain
            UseMxRecord                   = $connector.UseMxRecord
            CloudServicesMailEnabled      = $connector.CloudServicesMailEnabled
            RouteAllMessagesViaOnPremises = $connector.RouteAllMessagesViaOnPremises
            IsTransportRuleScoped         = $connector.IsTransportRuleScoped
            AllAcceptedDomains            = $connector.AllAcceptedDomains
            TestMode                      = $connector.TestMode
            ValidationRecipients          = ($connector.ValidationRecipients -join '; ')
            IsValidated                   = $connector.IsValidated
            LastValidationTimestamp       = $connector.LastValidationTimestamp
            SenderRewritingEnabled        = $connector.SenderRewritingEnabled
            LinkedConnector               = $connector.LinkedConnector
            Identity                      = $connector.Identity
            Guid                          = $connector.Guid
            WhenCreated                   = $connector.WhenCreated
            WhenChanged                   = $connector.WhenChanged
            DistinguishedName             = $connector.DistinguishedName
        }
    }
    
    # Display results summary
    Write-Host "`n==== OUTBOUND CONNECTORS SUMMARY ====" -ForegroundColor Cyan
    $results | Select-Object Name, Enabled, ConnectorType, TlsSettings, SmartHosts | Format-Table -AutoSize
    
    # Export to CSV
    Write-Host "`nExporting results to: $ExportPath" -ForegroundColor Cyan
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export completed successfully." -ForegroundColor Green
    
    # Summary statistics
    Write-Host "`n==== STATISTICS ====" -ForegroundColor Cyan
    Write-Host "Total Outbound Connectors: $($results.Count)" -ForegroundColor Green
    Write-Host "Enabled Connectors: $(($results | Where-Object {$_.Enabled -eq $true}).Count)" -ForegroundColor Green
    Write-Host "Disabled Connectors: $(($results | Where-Object {$_.Enabled -eq $false}).Count)" -ForegroundColor Yellow
    Write-Host "Validated Connectors: $(($results | Where-Object {$_.IsValidated -eq $true}).Count)" -ForegroundColor Green
    Write-Host "Test Mode Connectors: $(($results | Where-Object {$_.TestMode -eq $true}).Count)" -ForegroundColor Yellow
    
    Write-Host "`nConnector Types:" -ForegroundColor Yellow
    $results | Group-Object -Property ConnectorType | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor Gray
    }
    
    Write-Host "`nTLS Settings:" -ForegroundColor Yellow
    $results | Group-Object -Property TlsSettings | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor Gray
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
