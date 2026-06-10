<#
.SYNOPSIS
    Deploys the ACS Email Delivery Workbook and Portal Dashboard to Azure.

.DESCRIPTION
    Uses the Az PowerShell module to deploy:
      - workbook.json  : Azure Monitor Workbook (3-tab interactive analysis)
      - dashboard.json : Azure Portal Dashboard (pinnable KPI tiles + charts)

.PARAMETER TenantId
    Azure AD Tenant ID. Use this when you have access to multiple tenants to ensure the correct one is selected.

.PARAMETER SubscriptionId
    Azure Subscription ID where the workbook and dashboard will be deployed.

.PARAMETER ResourceGroupName
    Resource group to deploy the workbook and dashboard into.
    This can be the same RG as the Log Analytics workspace or a dedicated monitoring RG.

.PARAMETER WorkspaceName
    Name of the Log Analytics workspace that contains ACSEmailStatusUpdateOperational data.

.PARAMETER WorkspaceSubscriptionId
    Subscription ID where the Log Analytics workspace lives.
    Defaults to SubscriptionId if not specified (same subscription).

.PARAMETER WorkspaceResourceGroupName
    Resource group of the Log Analytics workspace.
    Defaults to ResourceGroupName if not specified.

.PARAMETER WorkbookDisplayName
    Display name for the workbook. Defaults to 'ACS Email Delivery Dashboard'.

.PARAMETER DashboardName
    Display name for the portal dashboard. Defaults to 'ACS Email Delivery Overview'.

.EXAMPLE
    # Deploy to same subscription as workspace
    .\deploy.ps1 `
        -SubscriptionId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ResourceGroupName "rg-monitoring" `
        -WorkspaceName "law-acs-prod"

    # Deploy to different subscription than workspace (cross-subscription)
    # Include TenantId if you have access to multiple Azure AD tenants
    .\deploy.ps1 `
        -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -SubscriptionId "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa" `
        -ResourceGroupName "rg-dashboards" `
        -WorkspaceName "law-acs-prod" `
        -WorkspaceSubscriptionId "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb" `
        -WorkspaceResourceGroupName "rg-logs"
#>
[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$SubscriptionId,

    [Parameter(Mandatory)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [string]$WorkspaceName,

    [Parameter()]
    [string]$WorkspaceSubscriptionId = $SubscriptionId,

    [Parameter()]
    [string]$WorkspaceResourceGroupName = $ResourceGroupName,

    [Parameter()]
    [string]$WorkbookDisplayName = "ACS Email Delivery Dashboard",

    [Parameter()]
    [string]$DashboardName = "ACS Email Delivery Overview"
)

$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot

# Ensure authenticated (handles multi-tenant/MFA scenarios)
$Account = Get-AzContext
if (-not $Account -or ($TenantId -and $Account.Tenant.Id -ne $TenantId)) {
    Write-Host "Authenticating to Azure..." -ForegroundColor Cyan
    $ConnectParams = @{}
    if ($TenantId) { $ConnectParams['TenantId'] = $TenantId }
    Connect-AzAccount @ConnectParams | Out-Null
}

Write-Host "Connecting to Azure subscription: $SubscriptionId" -ForegroundColor Cyan
$ContextParams = @{ SubscriptionId = $SubscriptionId }
if ($TenantId) {
    $ContextParams['TenantId'] = $TenantId
    Write-Host "Using tenant: $TenantId" -ForegroundColor Gray
}
Set-AzContext @ContextParams | Out-Null

$WorkspaceResourceId = "/subscriptions/$WorkspaceSubscriptionId/resourceGroups/$WorkspaceResourceGroupName/providers/Microsoft.OperationalInsights/workspaces/$WorkspaceName"

Write-Host "Log Analytics Workspace Resource ID: $WorkspaceResourceId" -ForegroundColor Gray

$WorkbookId   = [System.Guid]::NewGuid().ToString()
$DashboardId  = [System.Guid]::NewGuid().ToString()

Write-Host "`nDeploying Workbook..." -ForegroundColor Cyan
$WorkbookParams = @{
    ResourceGroupName                  = $ResourceGroupName
    TemplateFile                       = Join-Path $ScriptDir "workbook.json"
    workbookDisplayName                = $WorkbookDisplayName
    workbookId                         = $WorkbookId
    logAnalyticsWorkspaceResourceId    = $WorkspaceResourceId
}

$WorkbookDeployment = New-AzResourceGroupDeployment @WorkbookParams -Verbose
$WorkbookResourceId = $WorkbookDeployment.Outputs["workbookResourceId"].Value
Write-Host "Workbook deployed: $WorkbookResourceId" -ForegroundColor Green

Write-Host "`nDeploying Dashboard..." -ForegroundColor Cyan
$DashboardParams = @{
    ResourceGroupName                  = $ResourceGroupName
    TemplateFile                       = Join-Path $ScriptDir "dashboard.json"
    dashboardName                      = $DashboardName
    dashboardId                        = $DashboardId
    logAnalyticsWorkspaceResourceId    = $WorkspaceResourceId
    workbookResourceId                 = $WorkbookResourceId
}

$DashboardDeployment = New-AzResourceGroupDeployment @DashboardParams -Verbose
$DashboardResourceId = $DashboardDeployment.Outputs["dashboardResourceId"].Value
Write-Host "Dashboard deployed: $DashboardResourceId" -ForegroundColor Green

Write-Host "`n--- Deployment Complete ---" -ForegroundColor Green
Write-Host "Workbook  : https://portal.azure.com/#resource$WorkbookResourceId" -ForegroundColor Yellow
Write-Host "Dashboard : https://portal.azure.com/#resource$DashboardResourceId" -ForegroundColor Yellow
Write-Host "`nTo set the dashboard as your portal home page:" -ForegroundColor Gray
Write-Host "  Azure Portal -> Portal Settings -> My dashboards -> select '$DashboardName' -> Set as default" -ForegroundColor Gray
