$Result=@() 
$mailboxes = Get-Mailbox -ResultSize Unlimited
$totalmbx = $mailboxes.Count
$i = 1 
$mailboxes | ForEach-Object {
$i++
$mbx = $_
$mbs = Get-MailboxStatistics $mbx.UserPrincipalName
  
Write-Progress -activity "Processing $mbx" -status "$i out of $totalmbx completed"
  
if ($mbs.TotalItemSize -ne $null){
$size = [math]::Round(($mbs.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','')/1GB),2)
}else{
$size = 0 }
 
$Result += New-Object PSObject -property @{ 
Name = $mbx.DisplayName
UserPrincipalName = $mbx.UserPrincipalName
TotalSizeInMB = $size
SizeWarningQuota=$mbx.IssueWarningQuota
StorageSizeLimit = $mbx.ProhibitSendQuota
StorageLimitStatus = $mbs.ProhibitSendQuota
}
}
$Result | Export-CSV "C:\Export\MailboxSizeReport.csv" -NoTypeInformation -Encoding UNICODE