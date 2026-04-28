<#
    .SYNOPSIS
    Adds calendar permissions to user mailboxes for a mail-enabled security group with automatic language detection

    .DESCRIPTION
    Script is used to add calendar folder permissions to user mailboxes based on a mail-enabled security group.
    It automatically detects the calendar folder name regardless of the user's language (Danish, English, Italian, etc.)
    and logs all changes to a log file in the current directory.
    
    The script will automatically:
    - Check for Exchange Online Management module and install it if missing (works in non-admin context)
    - Connect to Exchange Online if not already connected
    - Detect calendar folder names in any language
    
    The permission level "Kan se titler og placeringer" (Can view titles and locations) corresponds to 
    LimitedDetails in Exchange.

    .PARAMETER Mailbox
    -Mailbox
        Specify the email address or identity of the user mailbox to add permissions to.
        Can be a single mailbox or multiple mailboxes (comma-separated or from pipeline).
        
        Required?                    true
        Position?                    0
        Default value
        Accept pipeline input?       true (ByValue, ByPropertyName)
        Accept wildcard characters?  false

    -SecurityGroup
        Specify the mail-enabled security group that will receive calendar permissions.
        
        Required?                    true
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -AccessRights
        Specify the permission level to grant. Default is "LimitedDetails" (Can view titles and locations).
        Valid values:
        - None: No access (Ingen)
        - AvailabilityOnly: Can see when I'm busy (Kan se, når jeg er optaget)
        - LimitedDetails: Can see titles and locations (Kan se titler og placeringer)
        - Reviewer: Can see all details (Kan se alle detaljer)
        - Editor: Can edit (Kan redigere)
        
        Required?                    false
        Position?                    2
        Default value                LimitedDetails
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -LogPath
        Specify optional custom path for the log file. Default is current directory with timestamp.
        
        Required?                    false
        Position?                    3
        Default value                .\CalendarPermissions_YYYYMMDD_HHMMSS.log
        Accept pipeline input?       false
        Accept wildcard characters?  false

    .EXAMPLE
    C:\PS> .\Add-CalendarPermissions.ps1 -Mailbox "user@domain.com" -SecurityGroup "My Organization"
    Adds LimitedDetails permission to user's calendar for "My Organization" group.

    .EXAMPLE
    C:\PS> .\Add-CalendarPermissions.ps1 -Mailbox "user@domain.com" -SecurityGroup "My Organization" -AccessRights Reviewer
    Adds Reviewer permission to user's calendar for "My Organization" group.

    .EXAMPLE
    C:\PS> Get-Mailbox -ResultSize Unlimited | .\Add-CalendarPermissions.ps1 -SecurityGroup "My Organization"
    Adds permissions to all mailboxes in the organization.

    .EXAMPLE
    C:\PS> .\Add-CalendarPermissions.ps1 -Mailbox "user@domain.com" -SecurityGroup "My Organization" -LogPath "C:\Logs\calendar.log"
    Adds permissions with custom log file path.

    .COPYRIGHT
    MIT License, feel free to distribute and use as you like, please leave author information.

    .LINK
    BLOG: http://www.hcandersen.net
    X: @dk_hcandersen

    .DISCLAIMER
    This script is provided AS-IS, with no warranty - Use at own risk.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('Identity','PrimarySmtpAddress','EmailAddress','TargetMailbox')]
    [string[]]$Mailbox,
    
    [Parameter(Mandatory=$true)]
    [string]$SecurityGroup,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('None','AvailabilityOnly','LimitedDetails','Reviewer','Editor')]
    [string]$AccessRights = "LimitedDetails",
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath
)

BEGIN {
    $ErrorActionPreference = "Continue"
    
    # Initialize log file
    if (-not $LogPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $LogPath = Join-Path -Path (Get-Location) -ChildPath "CalendarPermissions_$timestamp.log"
    }
    
    # Function to write to log
    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
            [string]$Level = 'INFO'
        )
        
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        
        # Write to log file
        Add-Content -Path $LogPath -Value $logMessage
        
        # Write to console with color
        switch ($Level) {
            'SUCCESS' { Write-Host $logMessage -ForegroundColor Green }
            'WARNING' { Write-Host $logMessage -ForegroundColor Yellow }
            'ERROR'   { Write-Host $logMessage -ForegroundColor Red }
            default   { Write-Host $logMessage -ForegroundColor White }
        }
    }
    
    # Function to check and install Exchange Online Management module
    function Install-ExchangeOnlineModule {
        Write-Log "Checking for Exchange Online Management module..." -Level INFO
        
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
        
        if (-not $module) {
            Write-Log "Exchange Online Management module not found. Installing..." -Level WARNING
            
            try {
                # Check if running as admin
                $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                
                if ($isAdmin) {
                    Write-Log "Installing module for all users (Admin context)..." -Level INFO
                    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
                } else {
                    Write-Log "Installing module for current user (Non-admin context)..." -Level INFO
                    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                }
                
                Write-Log "Exchange Online Management module installed successfully" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to install Exchange Online Management module: $($_.Exception.Message)" -Level ERROR
                throw "Module installation failed. Please install manually: Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser"
            }
        }
        else {
            Write-Log "Exchange Online Management module found (Version: $($module.Version))" -Level SUCCESS
        }
    }
    
    # Function to connect to Exchange Online
    function Connect-ExchangeOnlineService {
        Write-Log "Checking Exchange Online connection..." -Level INFO
        
        try {
            # Test if already connected by running a simple command
            $null = Get-OrganizationConfig -ErrorAction Stop
            Write-Log "Already connected to Exchange Online" -Level SUCCESS
            return
        }
        catch {
            Write-Log "Not connected to Exchange Online. Connecting..." -Level INFO
        }
        
        try {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            Write-Log "Successfully connected to Exchange Online" -Level SUCCESS
        }
        catch {
            Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level ERROR
            throw "Exchange Online connection failed. Please check your credentials and try again."
        }
    }
    
    # Function to get calendar folder path in any language
    function Get-CalendarFolderPath {
        param(
            [string]$MailboxIdentity
        )
        
        try {
            # Get all folders and find the calendar by FolderType
            $folders = Get-MailboxFolderStatistics -Identity $MailboxIdentity -FolderScope Calendar
            $calendarFolder = $folders | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
            
            if ($calendarFolder) {
                # Return the folder path with forward slashes replaced by backslashes
                $calendarPath = $calendarFolder.FolderPath -Replace '/', '\'
                return $calendarPath
            }
            else {
                throw "Could not detect calendar folder path"
            }
        }
        catch {
            throw "Error detecting calendar folder: $($_.Exception.Message)"
        }
    }
    
    # Start logging
    Write-Log "========================================" -Level INFO
    Write-Log "Calendar Permission Script Started" -Level INFO
    Write-Log "========================================" -Level INFO
    Write-Log "Security Group: $SecurityGroup" -Level INFO
    Write-Log "Access Rights: $AccessRights" -Level INFO
    Write-Log "Log File: $LogPath" -Level INFO
    Write-Log "========================================" -Level INFO
    
    # Install Exchange Online Management module if needed
    try {
        Install-ExchangeOnlineModule
    }
    catch {
        Write-Log "Critical error during module installation: $($_.Exception.Message)" -Level ERROR
        throw
    }
    
    # Connect to Exchange Online
    try {
        Connect-ExchangeOnlineService
    }
    catch {
        Write-Log "Critical error during Exchange Online connection: $($_.Exception.Message)" -Level ERROR
        throw
    }
    
    # Verify group exists (supports both distribution groups and mail-enabled security groups)
    try {
        # Try to get as recipient (works for both distribution groups and mail-enabled security groups)
        $group = Get-Recipient -Identity $SecurityGroup -ErrorAction Stop
        
        # Verify it's a group type
        if ($group.RecipientType -notlike "*Group*") {
            throw "The specified identity '$SecurityGroup' is not a group (Type: $($group.RecipientType))"
        }
        
        Write-Log "Group verified: $($group.DisplayName) ($($group.PrimarySmtpAddress))" -Level SUCCESS
        Write-Log "Group type: $($group.RecipientTypeDetails)" -Level INFO
        
        # Check if it's mail-enabled
        if (-not $group.PrimarySmtpAddress) {
            throw "Group '$SecurityGroup' is not mail-enabled"
        }
        
        # Get group members to add permissions individually
        Write-Log "Retrieving group members (permissions will be added for individual users)..." -Level INFO
        $groupMembers = Get-DistributionGroupMember -Identity $SecurityGroup -ResultSize Unlimited
        
        if ($groupMembers.Count -eq 0) {
            Write-Log "WARNING: Security group has no members!" -Level WARNING
            throw "Security group has no members to grant permissions to."
        }
        
        Write-Log "Found $($groupMembers.Count) member(s) in group" -Level SUCCESS
        
        # Filter to only user mailboxes (exclude contacts, groups, etc.)
        $script:groupUsers = $groupMembers | Where-Object { $_.RecipientType -like "*Mailbox" -or $_.RecipientType -eq "UserMailbox" }
        Write-Log "$($script:groupUsers.Count) user mailbox(es) will receive permissions" -Level INFO
        
        if ($script:groupUsers.Count -eq 0) {
            Write-Log "WARNING: No user mailboxes found in group!" -Level WARNING
            throw "Security group contains no user mailboxes."
        }
    }
    catch {
        Write-Log "ERROR: Failed to verify group '$SecurityGroup'" -Level ERROR
        Write-Log "Error details: $($_.Exception.Message)" -Level ERROR
        throw "Group validation failed"
    }
    
    $successCount = 0
    $errorCount = 0
    $skippedCount = 0
    $userPermissionCount = 0
}

PROCESS {
    foreach ($mbxIdentity in $Mailbox) {
        try {
            Write-Log "----------------------------------------" -Level INFO
            Write-Log "Processing mailbox: $mbxIdentity" -Level INFO
            
            # Get mailbox details
            $mbx = Get-EXOMailbox -Identity $mbxIdentity -ErrorAction Stop
            Write-Log "Mailbox found: $($mbx.DisplayName) ($($mbx.PrimarySmtpAddress))" -Level INFO
            
            # Detect calendar folder path
            Write-Log "Detecting calendar folder path..." -Level INFO
            $calendarFolderPath = Get-CalendarFolderPath -MailboxIdentity $mbx.PrimarySmtpAddress
            Write-Log "Calendar folder detected: $calendarFolderPath" -Level SUCCESS
            
            # Build calendar path
            $calendarPath = "$($mbx.PrimarySmtpAddress):$calendarFolderPath"
            
            # Add permissions for each user in the group
            $mailboxUserCount = 0
            $mailboxSkippedCount = 0
            
            foreach ($user in $script:groupUsers) {
                try {
                    $userIdentity = $user.PrimarySmtpAddress
                    
                    # Check if permission already exists for this user
                    $existingPermission = Get-MailboxFolderPermission -Identity $calendarPath -User $userIdentity -ErrorAction SilentlyContinue
                    
                    if ($existingPermission) {
                        if ($existingPermission.AccessRights -contains $AccessRights) {
                            Write-Log "  - $($user.DisplayName): Permission already exists. Skipping." -Level WARNING
                            $mailboxSkippedCount++
                            continue
                        }
                        else {
                            Write-Log "  - $($user.DisplayName): Updating from $($existingPermission.AccessRights -join ', ') to $AccessRights" -Level INFO
                            Set-MailboxFolderPermission -Identity $calendarPath -User $userIdentity -AccessRights $AccessRights -ErrorAction Stop
                            Write-Log "  - $($user.DisplayName): Permission updated successfully" -Level SUCCESS
                            $mailboxUserCount++
                        }
                    }
                    else {
                        Write-Log "  - $($user.DisplayName): Adding $AccessRights permission" -Level INFO
                        Add-MailboxFolderPermission -Identity $calendarPath -User $userIdentity -AccessRights $AccessRights -ErrorAction Stop
                        Write-Log "  - $($user.DisplayName): Permission added successfully" -Level SUCCESS
                        $mailboxUserCount++
                    }
                }
                catch {
                    Write-Log "  - $($user.DisplayName): ERROR - $($_.Exception.Message)" -Level ERROR
                }
            }
            
            Write-Log "Mailbox summary: $mailboxUserCount user(s) granted permissions, $mailboxSkippedCount skipped" -Level INFO
            $successCount++
            $userPermissionCount += $mailboxUserCount
            $skippedCount += $mailboxSkippedCount
        }
        catch {
            Write-Log "ERROR processing mailbox '$mbxIdentity': $($_.Exception.Message)" -Level ERROR
            $errorCount++
        }
    }
}

END {
    Write-Log "========================================" -Level INFO
    Write-Log "Calendar Permission Script Completed" -Level INFO
    Write-Log "========================================" -Level INFO
    Write-Log "Mailboxes Processed: $successCount" -Level SUCCESS
    Write-Log "Total User Permissions Granted: $userPermissionCount" -Level SUCCESS
    Write-Log "Total Errors: $errorCount" -Level ERROR
    Write-Log "Total Skipped: $skippedCount" -Level WARNING
    Write-Log "Log file saved to: $LogPath" -Level INFO
    Write-Log "========================================" -Level INFO
    
    # Display summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Mailboxes Processed: $successCount" -ForegroundColor Green
    Write-Host "User Permissions Granted: $userPermissionCount" -ForegroundColor Green
    Write-Host "Errors: $errorCount" -ForegroundColor Red
    Write-Host "Skipped: $skippedCount" -ForegroundColor Yellow
    Write-Host "Group Members: $($script:groupUsers.Count)" -ForegroundColor Cyan
    Write-Host "Log File: $LogPath" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}
