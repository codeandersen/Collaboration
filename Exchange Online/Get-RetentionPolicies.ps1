<#
.SYNOPSIS
Exports all Exchange Online user mailboxes with their retention policy.

.FIELDS
- PrimarySmtpAddress
- DisplayName
- RetentionPolicy

.OUTPUT
- Unicode CSV file
#>

# PARAMETERS
# Adjust the output path as needed
$OutputPath = ".\MailboxRetentionPolicies.csv"

# Ensure the Exchange Online module is available
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "ExchangeOnlineManagement module is not installed. Install it with: Install-Module ExchangeOnlineManagement"
    return
}

# Connect to Exchange Online (interactive login)
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

try {
    # Get all user mailboxes and select relevant properties
    $mailboxes = Get-mailbox -ResultSize Unlimited  | Select-Object PrimarySmtpAddress,DisplayName,RetentionPolicy

    # Export to Unicode CSV
    $mailboxes | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding Unicode

    Write-Host "Export complete. File saved to: $OutputPath"
}
finally {
    # Clean up session
    Disconnect-ExchangeOnline -Confirm:$false
}