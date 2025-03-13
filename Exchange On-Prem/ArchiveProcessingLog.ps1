$logProps = Export-MailboxDiagnosticLogs jg@mft-energy.com -ExtendedProperties

$xmlprops = [xml]($logProps.MailboxLog)

$xmlprops.Properties.MailboxTable.Property | ? {$_.Name -like "ELC*"}