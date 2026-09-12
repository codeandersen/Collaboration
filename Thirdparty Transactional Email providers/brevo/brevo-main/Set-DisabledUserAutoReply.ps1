# Script to set out of office messages for disabled user mailboxes
# Requires Exchange Online PowerShell Module and AzureAD Module

# Connect to Exchange Online using managed identity
Connect-ExchangeOnline -ManagedIdentity -Organization "toemmergaarden.onmicrosoft.com"

# Connect to Azure AD using managed identity
Connect-AzureAD -Identity

# Get all mailboxes
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { 
    $_.RecipientTypeDetails -eq "UserMailbox" 
}

foreach ($mailbox in $mailboxes) {
    try {
        # Check if user is disabled using Get-AzureADUser
        $user = Get-AzureADUser -ObjectId $mailbox.ExternalDirectoryObjectId
        if ($user.AccountEnabled -eq $false) {
            Write-Output "Setting auto-reply for disabled user: $($mailbox.UserPrincipalName)"

            # Set the auto-reply message
            $autoReplyMessage = "Jeg er stoppet. Mailen er sendt videre til x@tmg.dk"

            # Configure automatic replies
            Set-MailboxAutoReplyConfiguration -Identity $mailbox.UserPrincipalName `
                -AutoReplyState Enabled `
                -InternalMessage $autoReplyMessage `
                -ExternalMessage $autoReplyMessage `
                -ExternalAudience All

            Write-Output "Successfully set auto-reply for: $($mailbox.UserPrincipalName)"
        }
    }
    catch {
        Write-Error "Error processing mailbox $($mailbox.UserPrincipalName): $_"
    }
}

# Disconnect from services
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD
