## Variables for the script##
$PrimarySMTP = "jeba@scanvaegt.dk"

#Default CalendAr group
$DefaultCalendarGroup = ""
$DefaultCalendarGroupRights = "LimitedDetails"


#Author Celendar Group
$AuthorCalendarGroup = "ITsupport@scanvaegt.dk"
$AuthorCalendarGroupRights = "Editor"
## Variables for the script##

#Connect Exchange Online
try
{
    Write-Verbose "Logging in to Exchange Online..." -Verbose
    #Connect-ExchangeOnline
}

catch 
{   
    Write-Error -Message $_.Exception

    throw $_.Exception
}


Write-Verbose "Getting usermailbox with Primary SMTP: $PrimarySMTP"  -Verbose
#$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailBox | where-Object {($_.PrimarySMTPAddress -like "*@$PrimarySMTP")}
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailBox
Write-Verbose "$($Mailboxes.count) Users with PrimarySMTP: $PrimarySMTP was found"  -Verbose

$Count = 0

foreach ($Mailbox in $Mailboxes) {

    $Count++

    $Folder = Get-MailboxFolderStatistics -Identity $($Mailbox.UserPrincipalName) -FolderScope Calendar #-ErrorAction Stop

    Write-Verbose "Processing user $Count out of $($Mailboxes.count) Current user: $Mailbox" -Verbose
    Write-Verbose "Done Processing user: $Mailbox" -Verbose

    foreach ($F in $Folder) {

        if ($F.FolderType -eq 'Calendar') {

            $CalendarPath = $F.FolderPath -Replace '/', '\'

            #Set permissions for All users default group
            #Add-MailboxFolderPermission -Identity "$($Mailbox.UserPrincipalName):$CalendarPath" -User $DefaultCalendarGroup -AccessRights $DefaultCalendarGroupRights -ErrorAction SilentlyContinue

            #Set permissions for Author rights (Special Permissions)
            #Add-MailboxFolderPermission -Identity "$($Mailbox.UserPrincipalName):$CalendarPath" -User $AuthorCalendarGroup -AccessRights $AuthorCalendarGroupRights -ErrorAction SilentlyContinue
            Set-MailboxFolderPermission -Identity "$($Mailbox.UserPrincipalName):$CalendarPath" -User $AuthorCalendarGroup -AccessRights $AuthorCalendarGroupRights

            #Default "User" permissions
            #Set-MailboxFolderPermission -Identity "$($Mailbox.UserPrincipalName):$CalendarPath" -User Default -AccessRights LimitedDetails -ErrorAction SilentlyContinue
            
            Write-Verbose "Done Processing user: $Mailbox $CalendarPath" -Verbose

       }

    }

}