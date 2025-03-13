#connect
$Username = "XXXX"
$Password = ConvertTo-SecureString "XXXXX" -AsPlainText -Force
$LiveCred = New-Object System.Management.Automation.PSCredential $Username, $Password 

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $LiveCred -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking

#Ønsker ikke at se messagetrace på egne domæner (acceptedDomains)

$acceptedDomains = '@xx.dk|@xx.se'

# Startdato sættes til i går
$startdate = get-date -date $(get-date).adddays(-1) -format MM-dd-yyyy

#Slutdato til dagsdato
$enddate = get-date -format MM-dd-yyyy


#Collect Message Tracking Logs (These are broken into "pages" in Office 365 so we need to collect them all with a loop) 
$Messages = $null 
$Page = 1 
do 
{ 
Write-Host "Collecting Message Tracking - Page $Page..." 
$CurrMessages = Get-MessageTrace -StartDate $startdate -EndDate $enddate -pagesize 5000 -page $page -status "delivered" | Where { $_.RecipientAddress -notmatch $acceptedDomains -and $_.SenderAddress -notmatch 'XXX.onmicrosoft.com' } | select received, senderaddress, recipientaddress, messagetraceid

# Opretter en CSV for hver "side"

$CurrMessages | Export-Csv -notypeinformation -Path C:\XXX\GetMessageTrace_$page.csv
#$CurrMessages = Get-MessageTrace -PageSize 5000 -Page $Page | Select Received,SenderAddress,RecipientAddress,Size 
$Page++ 
$Messages += $CurrMessages 
} 

# Loop alle message traces igennem

until ($CurrMessages -eq $null) 



Write-Host "Now collecting MessageTrackingDetail for each Message..."


$Results = @()

foreach($trace in $Messages)
{

$Properties = @{

#### MessageTrace ###
RecipientAddress = $trace.recipientaddress
SenderAddress = $trace.senderaddress
MessageTraceID = $trace.messagetraceid
Subject = $trace.subject
Received = $trace.Received


###MessageTrace Detail ###

Detail = $msgdetail.detail
Event = $msgdetail.event
}



$Results += New-Object psobject -Property $properties

}

$Results | Select-Object received,messagetraceid, senderaddress, recipientaddress, event, detail | Export-Csv -notypeinformation -Path C:\xxx\MailFlow20181201.csv
