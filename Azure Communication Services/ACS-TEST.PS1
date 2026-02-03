$ACS_CLIENT_SECRET = ""
$Username = "app-acs-hr-starkgrp-prod"
$from = "noreply@HR.starkgroup.dk"
$to = "hans.christian.andersen@stark.dk"
$subject = "TEST ACS"
$body = "TEST ACS"


# Use your Entra app **client secret** as the password
$SecurePwd = ConvertTo-SecureString $ACS_CLIENT_SECRET -AsPlainText -Force  # or prompt/read safely
$Cred      = New-Object System.Management.Automation.PSCredential($Username, $SecurePwd)



Send-MailMessage -SmtpServer smtp.azurecomm.net -Credential $Cred -From $from -To $to -Subject $subject -Body $body -UseSsl