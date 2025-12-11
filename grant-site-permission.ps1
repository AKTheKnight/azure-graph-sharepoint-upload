<#+
Example usage:
    pwsh ./grant-site-permission.ps1 -TenantId "f0ac156c-04d2-4e80-a9fb-628a60467e33" -SiteId "aktheknight.sharepoint.com,8168cb8e-ec21-4d0a-9c50-a33a9980afa5,06e32acc-6158-4eaf-9533-9822ad6620f6" -AppClientId "3d5dbd42-507b-4cac-baeb-018cb458d00f" -Role write
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$SiteId,

    [Parameter(Mandatory = $true)]
    [string]$AppClientId,

    [ValidateSet("read", "write")]
    [string]$Role = "write"
)

Connect-MgGraph -TenantId $TenantId -Scopes "Sites.FullControl.All" -UseDeviceCode | Out-Null

$sp = Get-MgServicePrincipal -Filter "appId eq '$AppClientId'" | Select-Object -First 1
if (-not $sp) {
    throw "Unable to find service principal for app id $AppClientId"
}

$payload = @{
    roles = @($Role)
    grantedToIdentities = @(
        @{ application = @{ id = $sp.AppId; displayName = $sp.DisplayName } }
    )
} | ConvertTo-Json -Depth 4

Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions" -Body $payload -ContentType "application/json"
