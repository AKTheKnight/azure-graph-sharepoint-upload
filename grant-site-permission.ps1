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
