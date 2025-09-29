<#
.SYNOPSIS
    Ensures required service principal connections (Dataverse, SharePoint) exist in a target environment,
    and updates deploymentSettings.json with their connection IDs.

.PARAMETER EnvironmentUrl
    URL of the target Dataverse environment (e.g., https://org.crm.dynamics.com)

.PARAMETER TenantId
    Azure AD tenant ID

.PARAMETER ClientId
    Service principal app (client) ID

.PARAMETER ClientSecret
    Service principal client secret

.PARAMETER DeploymentSettingsPath
    Path to the deploymentSettings.json file to update
#>

param(
    [Parameter(Mandatory=$true)][string]$EnvironmentUrl,
    [Parameter(Mandatory=$true)][string]$TenantId,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [Parameter(Mandatory=$true)][string]$ClientSecret,
    [Parameter(Mandatory=$true)][string]$DeploymentSettingsPath
)

# --- Install modules if not already present ---
Write-Host "Installing Power Platform modules if missing..."
$modules = @('Microsoft.PowerApps.Administration.PowerShell', 'Microsoft.PowerApps.PowerShell')
foreach ($m in $modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Install-Module -Name $m -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module $m -Force
}

# --- Authenticate ---
Write-Host "Authenticating to Power Platform environment $EnvironmentUrl ..."
Add-PowerAppsAccount -Endpoint prod -TenantId $TenantId -ApplicationId $ClientId -ClientSecret $ClientSecret | Out-Null

# --- Helper function to get or create connection ---
function Ensure-Connection {
    param(
        [string]$ConnectorName,        # e.g. shared_commondataserviceforapps
        [string]$FriendlyName,         # For display
        [hashtable]$ConnectionParams   # Any additional params for creation
    )

    Write-Host "Checking for existing $FriendlyName connection..."
    $existing = Get-AdminPowerAppConnection -EnvironmentName $EnvironmentUrl | Where-Object { $_.ConnectorName -eq $ConnectorName }

    if ($existing) {
        Write-Host "$FriendlyName connection already exists: $($existing.ConnectionId)"
        return $existing.ConnectionId
    } else {
        Write-Host "Creating new $FriendlyName connection..."
        $newConn = New-AdminPowerAppConnection -ConnectorName $ConnectorName -ConnectionParameters $ConnectionParams -EnvironmentName $EnvironmentUrl
        Write-Host "$FriendlyName connection created: $($newConn.ConnectionId)"
        return $newConn.ConnectionId
    }
}

# --- Dataverse connection ---
$dataverseParams = @{
    "clientId"     = $ClientId
    "clientSecret" = $ClientSecret
    "tenantId"     = $TenantId
}
$dataverseConnectionId = Ensure-Connection -ConnectorName "shared_commondataserviceforapps" -FriendlyName "Dataverse" -ConnectionParams $dataverseParams

# --- SharePoint connection ---
$sharepointParams = @{
    "clientId"     = $ClientId
    "clientSecret" = $ClientSecret
    "tenantId"     = $TenantId
    # Optional: depending on how your SP connection is set up, you might need siteUrl or other fields
}
$sharepointConnectionId = Ensure-Connection -ConnectorName "shared_sharepointonline" -FriendlyName "SharePoint" -ConnectionParams $sharepointParams

# --- Update deploymentSettings.json ---
Write-Host "Patching deploymentSettings.json at $DeploymentSettingsPath ..."
$settings = Get-Content -Raw -Path $DeploymentSettingsPath | ConvertFrom-Json

foreach ($ref in $settings.ConnectionReferences) {
    if ($ref.LogicalName -like "*dataverse*") {
        $ref.ConnectionId = "/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps/connections/$dataverseConnectionId"
        $ref.ConnectorId  = "/providers/Microsoft.PowerApps/apis/shared_commondataserviceforapps"
    }
    elseif ($ref.LogicalName -like "*sharepoint*") {
        $ref.ConnectionId = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline/connections/$sharepointConnectionId"
        $ref.ConnectorId  = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
    }
}

$settings | ConvertTo-Json -Depth 10 | Set-Content -Path $DeploymentSettingsPath -Encoding UTF8
Write-Host "✅ Deployment settings updated successfully."
