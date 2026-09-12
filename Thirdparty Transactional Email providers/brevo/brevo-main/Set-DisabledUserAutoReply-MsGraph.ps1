# Script to set out of office messages for disabled user mailboxes
#
# Required Modules for Azure Automation:
# - Microsoft.Graph.Authentication
# - Microsoft.Graph.Users
# - Microsoft.Graph.Mail
#
# Required Permissions for Managed Identity:
# - User.Read.All
# - MailboxSettings.ReadWrite

# Connect to Microsoft Graph using managed identity
Connect-MgGraph -Identity

# Get all users that are disabled
$disabledUsers = Get-MgUser -All -Filter "accountEnabled eq false" -Property "id,userPrincipalName,mail"

foreach ($user in $disabledUsers) {
    try {
        # Check if user has Exchange Online mailbox
        $mailboxCheck = Get-MgUserMailboxSetting -UserId $user.Id -ErrorAction SilentlyContinue
        
        if ($mailboxCheck) {
            Write-Output "Setting auto-reply for disabled user: $($user.UserPrincipalName)"

            # Set the auto-reply message
            $autoReplyParams = @{
                AutomaticRepliesSetting = @{
                    Status = "AlwaysEnabled"
                    ExternalAudience = "All"
                    InternalReplyMessage = "Jeg er stoppet. Mailen er sendt videre til x@tmg.dk"
                    ExternalReplyMessage = "Jeg er stoppet. Mailen er sendt videre til x@tmg.dk"
                }
            }

            # Configure automatic replies using Microsoft Graph
            Update-MgUserMailboxSetting -UserId $user.Id -BodyParameter $autoReplyParams -WhatIf

            Write-Output "Successfully set auto-reply for: $($user.UserPrincipalName)"
        } else {
            Write-Output "Skipping user $($user.UserPrincipalName) - No Exchange Online mailbox found"
        }
    }
    catch {
        Write-Error "Error processing user $($user.UserPrincipalName): $_"
    }
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph