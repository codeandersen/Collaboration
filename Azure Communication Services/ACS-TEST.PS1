$ACS_CLIENT_SECRET = ""
$Username = ""

# Use your Entra app **client secret** as the password
$SecurePwd = ConvertTo-SecureString $ACS_CLIENT_SECRET -AsPlainText -Force  # or prompt/read safely
$Cred      = New-Object System.Management.Automation.PSCredential($Username, $SecurePwd)



Send-MailMessage -SmtpServer smtp.azurecomm.net -Credential $Cred -From noreply@test..at -To hca@xyz -Subject "TEST ACS" -Body "TEST ACS" -UseSsl