<# 
This script converts one or more Entra ID groups from 
on-premises–managed (AD-synced) to cloud-managed (SOA conversion).

- It uses Microsoft Graph PowerShell (`Invoke-MgGraphRequest`).
- Input: a list of group IDs.
- Output: table view in console + CSV log with results.

Requirements:
- Microsoft.Graph module installed
- Permissions: Group.ReadWrite.All (delegated)
- Role: Groups Administrator or Global Administrator
#>

# --- CONFIG ---
# Replace with your tenant ID
$TenantId = "xxx-xxx-xxx-xxx-xxx-xxx"

# Replace with one or more group IDs you want to convert
$GroupIds = @(
  "xxx-xxx-xxx-xxx-xxx",
  "xxx-xxx-xxx-xxx-xxx"
)

# --- AUTH ---
# Connect to Microsoft Graph with required scope in the right tenant
Connect-MgGraph -TenantId $TenantId -Scopes "Group.ReadWrite.All"

# --- FUNCTION ---
function Set-GroupSOACloudRest {
    param([string]$GroupId)

    # URI to change the group's Source of Authority
    $soaUri   = "https://graph.microsoft.com/beta/groups/$GroupId/onPremisesSyncBehavior"
    # URI to fetch the group's displayName (using $select to minimize payload)
    $grpUri   = "https://graph.microsoft.com/beta/groups/$GroupId" + '?$select=displayName,id'

    # Body to request SOA change -> make group cloud-managed
    $patchBody = @{ isCloudManaged = $true } | ConvertTo-Json

    try {
        # Step 1: Convert group to cloud-managed
        Invoke-MgGraphRequest -Method PATCH -Uri $soaUri -Body $patchBody -ContentType "application/json" | Out-Null

        # Small pause to let Graph commit the change
        Start-Sleep -Milliseconds 300

        # Step 2: Verify the SOA status
        $verify = Invoke-MgGraphRequest -Method GET -Uri $soaUri -ContentType "application/json"

        # Step 3: Get the group name for reporting
        $group  = Invoke-MgGraphRequest -Method GET -Uri $grpUri -ContentType "application/json"

        # Return structured result
        [pscustomobject]@{
            GroupId        = $GroupId
            DisplayName    = $group.displayName
            Status         = if ($verify.isCloudManaged) { "Converted" } else { "NotConverted" }
            IsCloudManaged = $verify.isCloudManaged
            Error          = $null
        }
    } catch {
        # Error handling -> capture failure and message
        [pscustomobject]@{
            GroupId        = $GroupId
            DisplayName    = $null
            Status         = "Failed"
            IsCloudManaged = $null
            Error          = $_.Exception.Message
        }
    }
}

# --- MAIN ---
# Loop over all group IDs, run the conversion function, collect results
$results = foreach ($id in $GroupIds) { Set-GroupSOACloudRest -GroupId $id }

# Display summary in console
$results | Format-Table GroupId,DisplayName,Status,IsCloudManaged

# Export full results (with error messages if any) to CSV
$results | Export-Csv -NoTypeInformation -Path ".\SOA-Conversion-Results.csv"

# Disconnect from Microsoft Graph
Disconnect-MgGraph