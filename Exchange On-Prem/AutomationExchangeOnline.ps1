Connect-AzureAD
$appId = "878e1146-68b8-4b23-b46b-983f92fa0a10"
$servicePrincipal = Get-AzureADServicePrincipal -Filter "AppID eq `'$appId'"
#Adding Role membership
$roleDefinition = Get-AzureADDirectoryRole | Where-Object {$_.DisplayName -eq 'Exchange Administrator'}
Add-AzureADDirectoryRoleMember -ObjectId $roleDefinition.ObjectId -RefObjectId $servicePrincipal.ObjectId
#Adding API Permissions
$api = (Get-AzureADServicePrincipal -Filter "AppID eq '00000002-0000-0ff1-ce00-000000000000'")
$permission = $api.AppRoles | Where-Object { $_.Value -eq 'Exchange.ManageAsApp' }
$apiPermission = [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]@{
    ResourceAppId  = $api.AppId ;
    ResourceAccess = [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
        Id   = $permission.Id ;
        Type = "Role"
    }
}
$Application = Get-AzureADApplication | Where-Object {$_.AppId -eq $appId}
$Application | Set-AzureADApplication -ReplyUrls 'http://localhost'
$Application | Set-AzureADApplication -RequiredResourceAccess $apiPermission