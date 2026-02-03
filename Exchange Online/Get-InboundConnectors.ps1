<#
.SYNOPSIS
    Gets all inbound connectors and their configuration from Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all inbound connectors
    with their complete configuration settings and exports them to a CSV file.

.PARAMETER ExportPath
    Path to export the results to a CSV file. If not specified, defaults to
    "InboundConnectors_<timestamp>.csv" in the current directory.

.EXAMPLE
    .\Get-InboundConnectors.ps1
    Exports inbound connectors to a timestamped CSV file in the current directory.

.EXAMPLE
    .\Get-InboundConnectors.ps1 -ExportPath "C:\Reports\InboundConnectors.csv"
    Exports inbound connectors to a specific CSV file.

.NOTES
    Requires: ExchangeOnlineManagement module
    Install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "InboundConnectors_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
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
    
    # Get all inbound connectors
    Write-Host "`nRetrieving all inbound connectors..." -ForegroundColor Cyan
    $inboundConnectors = Get-InboundConnector
    
    if ($inboundConnectors.Count -eq 0) {
        Write-Host "No inbound connectors found." -ForegroundColor Yellow
        exit 0
    }
    
    Write-Host "Found $($inboundConnectors.Count) inbound connector(s)." -ForegroundColor Green
    
    # Array to store results
    $results = @()
    
    # Process each inbound connector
    Write-Host "`nProcessing inbound connectors..." -ForegroundColor Cyan
    
    foreach ($connector in $inboundConnectors) {
        Write-Host "  Processing: $($connector.Name)..." -ForegroundColor Gray
        
        $results += [PSCustomObject]@{
            Name                          = $connector.Name
            Enabled                       = $connector.Enabled
            ConnectorType                 = $connector.ConnectorType
            ConnectorSource               = $connector.ConnectorSource
            Comment                       = $connector.Comment
            SenderDomains                 = ($connector.SenderDomains -join '; ')
            SenderIPAddresses             = ($connector.SenderIPAddresses -join '; ')
            TlsSenderCertificateName      = $connector.TlsSenderCertificateName
            RequireTls                    = $connector.RequireTls
            RestrictDomainsToCertificate  = $connector.RestrictDomainsToCertificate
            RestrictDomainsToIPAddresses  = $connector.RestrictDomainsToIPAddresses
            CloudServicesMailEnabled      = $connector.CloudServicesMailEnabled
            TreatMessagesAsInternal       = $connector.TreatMessagesAsInternal
            AssociatedAcceptedDomains     = ($connector.AssociatedAcceptedDomains -join '; ')
            EFSkipIPs                     = ($connector.EFSkipIPs -join '; ')
            EFSkipLastIP                  = $connector.EFSkipLastIP
            EFUsers                       = ($connector.EFUsers -join '; ')
            EFSkipMailGateway             = ($connector.EFSkipMailGateway -join '; ')
            ScanAndDropRecipients         = ($connector.ScanAndDropRecipients -join '; ')
            Identity                      = $connector.Identity
            Guid                          = $connector.Guid
            WhenCreated                   = $connector.WhenCreated
            WhenChanged                   = $connector.WhenChanged
            DistinguishedName             = $connector.DistinguishedName
        }
    }
    
    # Display results summary
    Write-Host "`n==== INBOUND CONNECTORS SUMMARY ====" -ForegroundColor Cyan
    $results | Select-Object Name, Enabled, ConnectorType, RequireTls, SenderDomains | Format-Table -AutoSize
    
    # Export to CSV
    Write-Host "`nExporting results to: $ExportPath" -ForegroundColor Cyan
    $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Export completed successfully." -ForegroundColor Green
    
    # Summary statistics
    Write-Host "`n==== STATISTICS ====" -ForegroundColor Cyan
    Write-Host "Total Inbound Connectors: $($results.Count)" -ForegroundColor Green
    Write-Host "Enabled Connectors: $(($results | Where-Object {$_.Enabled -eq $true}).Count)" -ForegroundColor Green
    Write-Host "Disabled Connectors: $(($results | Where-Object {$_.Enabled -eq $false}).Count)" -ForegroundColor Yellow
    Write-Host "TLS Required: $(($results | Where-Object {$_.RequireTls -eq $true}).Count)" -ForegroundColor Green
    
    Write-Host "`nConnector Types:" -ForegroundColor Yellow
    $results | Group-Object -Property ConnectorType | ForEach-Object {
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
