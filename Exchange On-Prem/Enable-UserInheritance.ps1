param(
    [Parameter(Mandatory = $true)]
    [string]$SearchBaseDN,

    [switch]$WhatIf
)

begin {
    # Import Active Directory module
    try {
        Import-Module ActiveDirectory
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to import ActiveDirectory module. Error: $($Error[0])"
        return
    }
}

process {
    Write-Host "Searching for user objects under: $SearchBaseDN"

    try {
        $users = Get-ADUser -LDAPFilter '(objectClass=user)' -SearchBase $SearchBaseDN -SearchScope Subtree -ErrorAction Stop
    }
    catch {
        Write-Host -ForegroundColor Red "Failed to search users under $SearchBaseDN. Error: $($Error[0])"
        return
    }

    if (-not $users) {
        Write-Host "No user objects found under $SearchBaseDN"
        return
    }

    foreach ($user in $users) {
        $dn = $user.DistinguishedName

        try {
            $acl = Get-Acl -Path "AD:\$dn"
        }
        catch {
            Write-Host -ForegroundColor Yellow "Could not read ACL for $dn. Skipping. Error: $($Error[0])"
            continue
        }

        # AreAccessRulesProtected = $true means inheritance is disabled
        if ($acl.AreAccessRulesProtected) {
            if ($WhatIf) {
                Write-Host "[WhatIf] Inheritance is disabled and would be enabled for user: $dn"
            }
            else {
                Write-Host "Enabling inheritance for user: $dn"
                try {
                    # Enable inheritance and preserve existing explicit ACEs
                    $acl.SetAccessRuleProtection($false, $true)
                    Set-Acl -Path "AD:\$dn" -AclObject $acl
                }
                catch {
                    Write-Host -ForegroundColor Red "Failed to enable inheritance for $dn. Error: $($Error[0])"
                }
            }
        }
    }
}

end {
    Write-Host -ForegroundColor Green "Enable-UserInheritance script completed."
}
