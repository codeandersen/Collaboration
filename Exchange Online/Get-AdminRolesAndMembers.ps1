<#
.SYNOPSIS
    Gets all admin roles and their members from Exchange Online.

.DESCRIPTION
    This script connects to Exchange Online and retrieves all administrative role groups
    and their members. The results can be exported to CSV or displayed on screen.

.PARAMETER ExportPath
    Optional path to export the results to a CSV file.

.EXAMPLE
    .\Get-AdminRolesAndMembers.ps1
    Displays all admin roles and members on screen.

.EXAMPLE
    .\Get-AdminRolesAndMembers.ps1 -ExportPath "C:\Reports\AdminRoles.csv"
    Exports admin roles and members to a CSV file.

.NOTES
    Requires: ExchangeOnlineManagement module
    Install: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath
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
    
    # Get all role groups
    Write-Host "`nRetrieving all admin role groups..." -ForegroundColor Cyan
    $roleGroups = Get-RoleGroup -ResultSize Unlimited
    
    Write-Host "Found $($roleGroups.Count) role groups." -ForegroundColor Green
    
    # Array to store results
    $results = @()
    
    # Process each role group
    Write-Host "`nProcessing role groups and members..." -ForegroundColor Cyan
    
    foreach ($roleGroup in $roleGroups) {
        Write-Host "  Processing: $($roleGroup.Name)..." -ForegroundColor Gray
        
        # Get members of the role group
        $members = Get-RoleGroupMember -Identity $roleGroup.Identity -ErrorAction SilentlyContinue
        
        if ($members) {
            foreach ($member in $members) {
                $results += [PSCustomObject]@{
                    RoleGroupName        = $roleGroup.Name
                    RoleGroupDescription = $roleGroup.Description
                    MemberName           = $member.Name
                    MemberDisplayName    = $member.DisplayName
                    MemberType           = $member.RecipientType
                    MemberEmail          = $member.PrimarySmtpAddress
                    RoleGroupGuid        = $roleGroup.Guid
                    MemberGuid           = $member.Guid
                }
            }
        }
        else {
            # Add role group with no members
            $results += [PSCustomObject]@{
                RoleGroupName        = $roleGroup.Name
                RoleGroupDescription = $roleGroup.Description
                MemberName           = "No Members"
                MemberDisplayName    = "No Members"
                MemberType           = "N/A"
                MemberEmail          = "N/A"
                RoleGroupGuid        = $roleGroup.Guid
                MemberGuid           = "N/A"
            }
        }
    }
    
    # Display results
    Write-Host "`n==== ADMIN ROLES AND MEMBERS ====" -ForegroundColor Cyan
    Write-Host "Total Records: $($results.Count)" -ForegroundColor Green
    $results | Format-Table -AutoSize
    
    # Export to CSV if path is provided
    if ($ExportPath) {
        Write-Host "`nExporting results to: $ExportPath" -ForegroundColor Cyan
        $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Export completed successfully." -ForegroundColor Green
    }
    
    # Summary statistics
    Write-Host "`n==== SUMMARY ====" -ForegroundColor Cyan
    Write-Host "Total Role Groups: $($roleGroups.Count)" -ForegroundColor Green
    Write-Host "Total Role Assignments: $($results.Count)" -ForegroundColor Green
    
    $groupedByRole = $results | Group-Object -Property RoleGroupName
    Write-Host "`nRole Groups with Most Members:" -ForegroundColor Yellow
    $groupedByRole | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count) member(s)" -ForegroundColor Gray
    }
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
