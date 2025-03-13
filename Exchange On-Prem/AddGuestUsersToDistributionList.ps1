#$connection = Get-AutomationConnection –Name AzureRunAsConnection
$tenant = 'mftenergy.onmicrosoft.com'
$certThumbprint = "5e64c50b590f755b8dd96e00ade69506254b4c8d"
$appId = "878e1146-68b8-4b23-b46b-983f92fa0a10"

function ConnectExchangeOnlineWithCert () {
	try
	{
    	"Logging in to Exchange Online..."
    	#Connect-ExchangeOnline –CertificateThumbprint $connection.CertificateThumbprint –AppId $connection.ApplicationID –ShowBanner:$false –Organization $tenant
    	Connect-ExchangeOnline -CertificateThumbprint $certThumbprint -AppId $appId -Organization $tenant –ShowBanner:$false 
	}catch {   
    Write-Error -Message $_.Exception
    throw $_.Exception
	}
}

function ConnectAzureADAsMSI () {
Connect-azaccount -identity
$context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
$graphToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, "https://graph.microsoft.com").AccessToken
$aadToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, "https://graph.windows.net").AccessToken
	Try {
		Connect-AzureAD -AadAccessToken $aadToken -AccountId $context.Account.Id -TenantId $context.tenant.id
		Write-Verbose "AzureAD Connection established!" -Verbose
		} catch {
		Write-Error $_
		}
}

ConnectExchangeOnlineWithCert
ConnectAzureADAsMSI

#Do some ExchangeOnline and AzureAD scripting here!

$Group = "# MFT Guest Account"
$GuestUsers = Get-AzureADUser -All $true | Where-Object {$_.UserType -eq 'Guest'}
#Get-AzureADUser -All $true | Where-Object {$_.UserType -eq 'Guest'} | Select UserPrincipalName
Try {
	ForEach ($User in $GuestUsers){
    	If (-not(Get-DistributionGroupMember $Group | Where-Object WindowsLiveID -EQ $User.UserPrincipalName))
        	{
            	Add-DistributionGroupMember $Group -Member $User.UserPrincipalName
				
            	Write-Output "User" $User.UserPrincipalName "added to distribution list" $Group -Verbose
        	}
	}
} catch {
	Write-Error $_
}

Disconnect-ExchangeOnline -Confirm:$false

