<#
        .SYNOPSIS
        Creates Microsoft 365 Group from parameters, adds folder to document library and adds security group as member

        .DESCRIPTION
        Script is used to create an Microsoft 365 Group with Sharepoint document library, add folders and give access to document library through a security group.

        .PARAMETER Name
        -GroupName
            Specify the name of the group.
            
            Required?                    true
               Position?                    0
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        -Folders
            Specify the name of the folders to be created in the sharepoint site.
            
            Required?                    true
               Position?                    0
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        -SecurityGroup
            Specify the name of the security to be added as a member.
            
            Required?                    true
               Position?                    0
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        .PARAMETER Extension
        Specifies the extension. "GroupSMTPAddress" is the default.

               Required?                    True
               Position?                    1
               Default value
               Accept pipeline input?       false
               Accept wildcard characters?

        .EXAMPLE
        C:\PS> Create-MS365Group.ps1 -GroupName "Finance Department" -GroupSMTPAddress Finance@contoso.onmicrosoft.com -Folders Finance,Finance_External -SecurityGroup sg-financedepartment

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

       .LINK
        BLOG: http://www.hcconsult.dk
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
    #>
    Param(
        [Parameter(Mandatory=$True,
        Position=0,
        HelpMessage="Enter name of group to be created")]
        [String]$Groupname
        ,
        [Parameter(Mandatory=$True,
        Position=1,
        HelpMessage="Enter name/names of folders to be created in Sharepoint Site")]
        [String]$Folders
        ,
        [Parameter(Mandatory=$True,
        Position=2,
        HelpMessage="Enter the name of the security group")]
        [String]$SecurityGroup
        )
        
        

Function ConnectExchangeOnline () {
    #Connect to Exchange Online PowerShell
    Write-Verbose "Connecting to Exchange Online" -Verbose
    
        if(!(Get-Module ExchangeOnlineManagement)){
            Install-Module -Name ExchangeOnlineManagement -Force
        }
    
    Try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Verbose "Connecting to Exchange Online completed." -Verbose
    }catch{
        Throw "Could not connect to Exchange Online, make sure to get the exchange online v2 module installed. $_"
    }
}

Function ConnectAzureAD () {
    #Connect to Azure AD PowerShell
    Write-Verbose "Connecting to Azure AD" -Verbose
    
        if(!(Get-Module AzureAD)){
            Install-Module -Name AzureAD -Force
        }
    
    Try {
        Connect-AzureAD -ErrorAction Stop
        Write-Verbose "Connecting to Azure AD completed." -Verbose
    }catch{
        Throw "Could not connect to Azure AD, make sure to get the Azure AD module installed"
    }
}

Function ConnectSharepointOnline ($strSharepointUrl) {
    #Connect to Sharepoint Online PowerShell
    Write-Verbose "Connecting to Sharepoint Online" -Verbose
    
        if(!(Get-Module PnP.PowerShell)){
            Install-Module -Name PnP.PowerShell -Force
        }
    
    Try {
        Connect-PnPOnline -Url $strSharepointUrl -UseWebLogin -ReturnConnection    
        Write-Verbose "Connecting to Sharepoint Online completed." -Verbose 
    }catch{
        Throw "Could not connect to Sharepoint Online, make sure to get the Sharepoint Online module installed"
    }
}


Function  CreateMS365Group ($StrGroupname) {
    #Create MS365 and return Sharepoint Document URL
    Write-Verbose "Creating Microsoft 365 Group"
    If (!(Get-UnifiedGroup | Where-Object DisplayName -eq "$StrGroupname"))
        {
        New-AzureADMSGroup -DisplayName "$StrGroupname" -GroupTypes Unified -MailEnabled $True -MailNickname "$StrGroupname" -SecurityEnabled $True -Visibility Private  
        #Get SharepointUrl
        Do {	
            $SharePointSiteUrl = Get-UnifiedGroup | Where-Object Displayname -eq "$StrGroupname" | Select-Object -ExpandProperty SharePointSiteUrl
            If ( [string]::IsNullOrWhiteSpace($SharePointSiteUrl)) {
                $SharePointSiteUrl = $false
                Write-host "Sharepoint Document library not created yet"
            }
            Else {
                Write-host "Sharepoint Document library created!"
                $SharePoint = Get-UnifiedGroup | Where-Object Displayname -eq "$GroupName" | Select-Object -ExpandProperty SharePointSiteUrl
                Return $SharePoint      
            }             
        }
        While (
            $SharePointSiteUrl -eq $false
            )    
        }
    }


Function CreateSharepointOnlineFolders ($strSharepointUrl,$strFolders) {
    Write-Verbose "Adding folders to Sharepoint Document library of Microsoft 365 Group"
    If ($strFolders)
        {
        $connection = Connect-PnPOnline -Url $strSharepointUrl -UseWebLogin -ReturnConnection
        $StrFoldersList = $strFolders.Split(",")
        ForEach ($strFolder in $StrFoldersList) 
        {
            Write-Verbose "Adding folder $strFolder to Sharepoint Document"
            Resolve-PnPFolder -SiteRelativePath "/Shared Documents/$strFolder" -Connection $connection  
        }   
        
        }
    Else {
        Write-Output "Folders missing in input variable strFolders"
    }
    }        

    Function AddGroupToSharepointOnlineSite ($strSharepointUrl,$strGroupname,$strSecurityGroup) {
        Write-Verbose "Adding security group to sharepoint members of Sharepoint Site"
        If ($strSharepointUrl)
            {
            $ADGroupID = Get-AzureADGroup -SearchString "$strSecurityGroup" | Select-Object -ExpandProperty ObjectId
            $connection = Connect-PnPOnline -Url $strSharepointUrl -UseWebLogin -ReturnConnection
            $LoginName = "c:0t`.c`|tenant`|$ADGroupID"       
            Add-PnPGroupMember -LoginName $LoginName -Identity "$strGroupname Members" -Connection $connection
            }
        Else {
            Write-Output "Sharepoint URL missing in input variable strSharepointUrl"
        }
        }        

ConnectExchangeOnline 
ConnectAzureAD
$SharepointUrl = CreateMS365Group $Groupname | Select-Object -Skip 1
ConnectSharepointOnline $SharepointUrl
CreateSharepointOnlineFolders $SharepointUrl $Folders 
AddGroupToSharepointOnlineSite $SharepointUrl $Groupname $SecurityGroup
$StrFoldersList = $Folders.Split(",")
ForEach ($strFolder in $StrFoldersList) 
        {
            Write-Output "Script completed Sharepoint Doc URLs: $SharepointUrl/Shared Documents/$strFolder"            
        }   

#Exit
Get-PSSession | Remove-PSSession
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD -Confirm:$false
