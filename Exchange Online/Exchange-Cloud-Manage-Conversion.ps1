#Requires -Version 5.1

<#
.SYNOPSIS
    Exchange Cloud Manage Conversion Tool
.DESCRIPTION
    GUI tool to manage Exchange mailbox cloud/on-premises management attribute
.NOTES
    Author: Exchange Management Tool
    Version: 1.0
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogFile = Join-Path $script:ScriptPath "ExchangeCloudManagement_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8
    
    Write-Host $logEntry
}

function Search-ExchangeOnlineModule {
    Write-Log "Checking for Exchange Online Management module..."
    
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    
    if (-not $module) {
        Write-Log "Exchange Online Management module not found. Attempting to install..." -Level WARNING
        
        try {
            [System.Windows.Forms.MessageBox]::Show(
                "Exchange Online Management module is not installed.`n`nThe tool will now attempt to install it. This may take a few minutes.",
                "Module Installation Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Log "Exchange Online Management module installed successfully." -Level INFO
            
            [System.Windows.Forms.MessageBox]::Show(
                "Exchange Online Management module has been installed successfully.",
                "Installation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            return $true
        }
        catch {
            Write-Log "Failed to install Exchange Online Management module: $($_.Exception.Message)" -Level ERROR
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to install Exchange Online Management module.`n`nError: $($_.Exception.Message)`n`nPlease install manually using:`nInstall-Module -Name ExchangeOnlineManagement -Scope CurrentUser",
                "Installation Failed",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            
            return $false
        }
    }
    else {
        Write-Log "Exchange Online Management module is already installed (Version: $($module.Version))."
        return $true
    }
}

function Connect-ExchangeOnlineSession {
    Write-Log "Attempting to connect to Exchange Online..."
    
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        
        Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false
        
        Write-Log "Successfully connected to Exchange Online."
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" -Level ERROR
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Exchange Online.`n`nError: $($_.Exception.Message)",
            "Connection Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $false
    }
}

function Get-ExchangeUsers {
    Write-Log "Retrieving mailbox users from Exchange Online..."
    
    try {
        $allMailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop | Select-Object DisplayName, PrimarySmtpAddress, IsExchangeCloudManaged, UserPrincipalName, IsDirSynced, RecipientTypeDetails
        
        Write-Log "Retrieved $($allMailboxes.Count) total mailbox users."
        
        $hybridMailboxes = $allMailboxes | Where-Object { $_.IsDirSynced -eq $true }
        
        $cloudOnlyCount = $allMailboxes.Count - $hybridMailboxes.Count
        Write-Log "Filtered out $cloudOnlyCount cloud-only mailboxes. Displaying $($hybridMailboxes.Count) hybrid/on-premises synced mailboxes."
        
        return $hybridMailboxes
    }
    catch {
        Write-Log "Failed to retrieve mailbox users: $($_.Exception.Message)" -Level ERROR
        
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to retrieve mailbox users.`n`nError: $($_.Exception.Message)",
            "Retrieval Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        
        return $null
    }
}

function Convert-ToCloudManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedUser
    )
    
    $identity = $SelectedUser.UserPrincipalName
    $displayName = $SelectedUser.DisplayName
    
    Write-Log "Converting user '$displayName' ($identity) to Cloud Managed..."
    
    try {
        Set-Mailbox -Identity $identity -IsExchangeCloudManaged $true -ErrorAction Stop
        
        Write-Log "Successfully converted user '$displayName' ($identity) to Cloud Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to convert user '$displayName' ($identity) to Cloud Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

function Convert-ToOnPremManaged {
    param(
        [Parameter(Mandatory=$true)]
        $SelectedUser
    )
    
    $identity = $SelectedUser.UserPrincipalName
    $displayName = $SelectedUser.DisplayName
    
    Write-Log "Converting user '$displayName' ($identity) to On-Premises Managed..."
    
    try {
        Set-Mailbox -Identity $identity -IsExchangeCloudManaged $false -ErrorAction Stop
        
        Write-Log "Successfully converted user '$displayName' ($identity) to On-Premises Managed." -Level INFO
        
        return $true
    }
    catch {
        Write-Log "Failed to convert user '$displayName' ($identity) to On-Premises Managed. Error: $($_.Exception.Message)" -Level ERROR
        
        return $false
    }
}

Write-Log "========================================" -Level INFO
Write-Log "Exchange Cloud Manage Conversion Tool Started" -Level INFO
Write-Log "========================================" -Level INFO

if (-not (Search-ExchangeOnlineModule)) {
    Write-Log "Cannot proceed without Exchange Online Management module. Exiting." -Level ERROR
    exit 1
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange Cloud Manage Conversion"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Location = New-Object System.Drawing.Point(10, 10)
$labelTitle.Size = New-Object System.Drawing.Size(880, 30)
$labelTitle.Text = "Exchange Cloud Management Conversion Tool"
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$labelTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($labelTitle)

$buttonConnect = New-Object System.Windows.Forms.Button
$buttonConnect.Location = New-Object System.Drawing.Point(10, 50)
$buttonConnect.Size = New-Object System.Drawing.Size(150, 30)
$buttonConnect.Text = "Connect to EXO"
$buttonConnect.Add_Click({
    $buttonConnect.Enabled = $false
    $buttonRefresh.Enabled = $false
    
    if (Connect-ExchangeOnlineSession) {
        $buttonConnect.Text = "Connected"
        $buttonConnect.Enabled = $false
        $buttonRefresh.Enabled = $true
        $buttonDisconnect.Enabled = $true
        
        $users = Get-ExchangeUsers
        
        if ($users) {
            $dataGridView.Rows.Clear()
            
            foreach ($user in $users) {
                $cloudManagedStatus = if ($user.IsExchangeCloudManaged) { "True" } else { "False" }
                $dataGridView.Rows.Add($user.DisplayName, $user.PrimarySmtpAddress, $cloudManagedStatus, $user.UserPrincipalName)
            }
            
            $statusLabel.Text = "Status: Connected - $($users.Count) users loaded"
        }
    }
    else {
        $buttonConnect.Enabled = $true
    }
})
$form.Controls.Add($buttonConnect)

$buttonRefresh = New-Object System.Windows.Forms.Button
$buttonRefresh.Location = New-Object System.Drawing.Point(170, 50)
$buttonRefresh.Size = New-Object System.Drawing.Size(100, 30)
$buttonRefresh.Text = "Refresh Users"
$buttonRefresh.Enabled = $false
$buttonRefresh.Add_Click({
    $buttonRefresh.Enabled = $false
    
    $users = Get-ExchangeUsers
    
    if ($users) {
        $dataGridView.Rows.Clear()
        
        foreach ($user in $users) {
            $cloudManagedStatus = if ($user.IsExchangeCloudManaged) { "True" } else { "False" }
            $dataGridView.Rows.Add($user.DisplayName, $user.PrimarySmtpAddress, $cloudManagedStatus, $user.UserPrincipalName)
        }
        
        $statusLabel.Text = "Status: Refreshed - $($users.Count) users loaded"
    }
    
    $buttonRefresh.Enabled = $true
})
$form.Controls.Add($buttonRefresh)

$buttonDisconnect = New-Object System.Windows.Forms.Button
$buttonDisconnect.Location = New-Object System.Drawing.Point(280, 50)
$buttonDisconnect.Size = New-Object System.Drawing.Size(150, 30)
$buttonDisconnect.Text = "Disconnect from EXO"
$buttonDisconnect.Enabled = $false
$buttonDisconnect.Add_Click({
    Write-Log "Disconnecting from Exchange Online..."
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Log "Successfully disconnected from Exchange Online."
        
        $buttonConnect.Text = "Connect to EXO"
        $buttonConnect.Enabled = $true
        $buttonRefresh.Enabled = $false
        $buttonDisconnect.Enabled = $false
        
        $dataGridView.Rows.Clear()
        
        $statusLabel.Text = "Status: Disconnected. Click 'Connect to EXO' to begin."
        
        [System.Windows.Forms.MessageBox]::Show(
            "Successfully disconnected from Exchange Online.",
            "Disconnected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
        
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred while disconnecting.`n`nError: $($_.Exception.Message)",
            "Disconnect Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})
$form.Controls.Add($buttonDisconnect)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 90)
$dataGridView.Size = New-Object System.Drawing.Size(870, 380)
$dataGridView.AllowUserToAddRows = $false
$dataGridView.AllowUserToDeleteRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $true
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill

$colDisplayName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDisplayName.Name = "DisplayName"
$colDisplayName.HeaderText = "Display Name"
$colDisplayName.FillWeight = 30
[void]$dataGridView.Columns.Add($colDisplayName)

$colEmail = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colEmail.Name = "Email"
$colEmail.HeaderText = "Email Address"
$colEmail.FillWeight = 35
[void]$dataGridView.Columns.Add($colEmail)

$colCloudManaged = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colCloudManaged.Name = "IsExchangeCloudManaged"
$colCloudManaged.HeaderText = "Is Exchange Cloud Managed"
$colCloudManaged.FillWeight = 25
[void]$dataGridView.Columns.Add($colCloudManaged)

$colUserPrincipalName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUserPrincipalName.Name = "UserPrincipalName"
$colUserPrincipalName.HeaderText = "User Principal Name"
$colUserPrincipalName.Visible = $false
[void]$dataGridView.Columns.Add($colUserPrincipalName)

$form.Controls.Add($dataGridView)

$buttonConvertToCloud = New-Object System.Windows.Forms.Button
$buttonConvertToCloud.Location = New-Object System.Drawing.Point(10, 480)
$buttonConvertToCloud.Size = New-Object System.Drawing.Size(200, 35)
$buttonConvertToCloud.Text = "Convert to Cloud Managed"
$buttonConvertToCloud.BackColor = [System.Drawing.Color]::LightGreen
$buttonConvertToCloud.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one user from the list.",
            "No User Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedCount = $dataGridView.SelectedRows.Count
    $userList = ($dataGridView.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to convert user '$userList' to Cloud Managed?"
    } else {
        "Are you sure you want to convert $selectedCount users to Cloud Managed?`n`nUsers: $userList"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $selectedUser = @{
                DisplayName = $selectedRow.Cells["DisplayName"].Value
                PrimarySmtpAddress = $selectedRow.Cells["Email"].Value
                UserPrincipalName = $selectedRow.Cells["UserPrincipalName"].Value
            }
            
            if (Convert-ToCloudManaged -SelectedUser $selectedUser) {
                $selectedRow.Cells["IsExchangeCloudManaged"].Value = "True"
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed: $failCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Conversion Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        Write-Log "Batch conversion to Cloud Managed completed. Success: $successCount, Failed: $failCount"
    }
})
$form.Controls.Add($buttonConvertToCloud)

$buttonConvertToOnPrem = New-Object System.Windows.Forms.Button
$buttonConvertToOnPrem.Location = New-Object System.Drawing.Point(220, 480)
$buttonConvertToOnPrem.Size = New-Object System.Drawing.Size(200, 35)
$buttonConvertToOnPrem.Text = "Convert to On-Prem Managed"
$buttonConvertToOnPrem.BackColor = [System.Drawing.Color]::LightCoral
$buttonConvertToOnPrem.Add_Click({
    if ($dataGridView.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select at least one user from the list.",
            "No User Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $selectedCount = $dataGridView.SelectedRows.Count
    $userList = ($dataGridView.SelectedRows | ForEach-Object { $_.Cells["DisplayName"].Value }) -join ", "
    
    $confirmMessage = if ($selectedCount -eq 1) {
        "Are you sure you want to convert user '$userList' to On-Premises Managed?"
    } else {
        "Are you sure you want to convert $selectedCount users to On-Premises Managed?`n`nUsers: $userList"
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm Conversion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $successCount = 0
        $failCount = 0
        
        foreach ($selectedRow in $dataGridView.SelectedRows) {
            $selectedUser = @{
                DisplayName = $selectedRow.Cells["DisplayName"].Value
                PrimarySmtpAddress = $selectedRow.Cells["Email"].Value
                UserPrincipalName = $selectedRow.Cells["UserPrincipalName"].Value
            }
            
            if (Convert-ToOnPremManaged -SelectedUser $selectedUser) {
                $selectedRow.Cells["IsExchangeCloudManaged"].Value = "False"
                $successCount++
            } else {
                $failCount++
            }
        }
        
        $summaryMessage = "Batch conversion completed.`n`nSuccessful: $successCount`nFailed: $failCount"
        
        [System.Windows.Forms.MessageBox]::Show(
            $summaryMessage,
            "Batch Conversion Summary",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        Write-Log "Batch conversion to On-Premises Managed completed. Success: $successCount, Failed: $failCount"
    }
})
$form.Controls.Add($buttonConvertToOnPrem)

$buttonOpenLog = New-Object System.Windows.Forms.Button
$buttonOpenLog.Location = New-Object System.Drawing.Point(430, 480)
$buttonOpenLog.Size = New-Object System.Drawing.Size(150, 35)
$buttonOpenLog.Text = "Open Log File"
$buttonOpenLog.BackColor = [System.Drawing.Color]::LightBlue
$buttonOpenLog.Add_Click({
    if (Test-Path $script:LogFile) {
        try {
            Start-Process notepad.exe -ArgumentList $script:LogFile
            Write-Log "Log file opened: $script:LogFile"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to open log file.`n`nError: $($_.Exception.Message)",
                "Error Opening Log",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "Log file does not exist yet.`n`nPath: $script:LogFile",
            "Log File Not Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})
$form.Controls.Add($buttonOpenLog)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 525)
$statusLabel.Size = New-Object System.Drawing.Size(870, 20)
$statusLabel.Text = "Status: Not connected. Click 'Connect to Exchange Online' to begin."
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$form.Controls.Add($statusLabel)

$form.Add_FormClosing({
    Write-Log "Exchange Cloud Manage Conversion Tool Closing" -Level INFO
    
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Exchange Online." -Level INFO
    }
    catch {
        Write-Log "Error during disconnect: $($_.Exception.Message)" -Level WARNING
    }
})

[void]$form.ShowDialog()

Write-Log "========================================" -Level INFO
Write-Log "Exchange Cloud Manage Conversion Tool Ended" -Level INFO
Write-Log "========================================" -Level INFO
