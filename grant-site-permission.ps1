#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites
<#
.SYNOPSIS
    Grants Microsoft Graph Sites.Selected access for an app or managed identity to a specific SharePoint site.
.DESCRIPTION
    Uses the Microsoft.Graph PowerShell SDK to connect with delegated Sites.FullControl.All scope, resolves
    the site (by id or name), and posts the permission grant to /sites/{id}/permissions.
    Requires admin consent for Sites.Selected on the target principal and Sites.FullControl.All for the operator.
.EXAMPLE
    pwsh ./grant-site-permission.ps1 -SiteUrl "https://aktheknight.sharepoint.com/sites/testsite" -PrincipalAppId 3d5dbd42-507b-4cac-baeb-018cb458d00f -PrincipalDisplayName "graph-upload-local" -Role write -TenantId f0ac156c-04d2-4e80-a9fb-628a60467e33 -ExtraScopes "Application.Read.All"
#>
[CmdletBinding(DefaultParameterSetName = "ByName")]
param(
    [Parameter(ParameterSetName = "ById", Mandatory = $true)]
    [string]$SiteId,

    [Parameter(ParameterSetName = "ByName", Mandatory = $true)]
    [string]$SiteName,

    [Parameter(ParameterSetName = "ByUrl", Mandatory = $true)]
    [string]$SiteUrl,

    [ValidateSet("read", "write")]
    [string]$Role = "write",

    [string]$PrincipalObjectId,
    [string]$PrincipalAppId,
    [string]$PrincipalDisplayName,

    [string]$TenantId,
    [string[]]$ExtraScopes = @()
)

if (-not $PrincipalObjectId -and -not $PrincipalAppId) {
    throw "Provide either -PrincipalObjectId or -PrincipalAppId (clientId)."
}

$requiredScopes = @("Sites.FullControl.All") + $ExtraScopes

function Ensure-GraphConnection {
    $context = Get-MgContext
    if ($TenantId -and $context -and $context.TenantId -ne $TenantId) {
        Disconnect-MgGraph | Out-Null
        $context = $null
    }

    $hasScope = $false
    if ($context -and $context.Scopes) {
        $contextScopes = @($context.Scopes)
        $matchCount = ($requiredScopes | Where-Object { $contextScopes -contains $_ }).Count
        $hasScope = $matchCount -eq $requiredScopes.Count
    }

    if (-not $context -or -not $hasScope) {
        Write-Host "Connecting to Microsoft Graph using device code..." -ForegroundColor Cyan
        if ($TenantId) {
            Connect-MgGraph -Scopes $requiredScopes -UseDeviceCode -TenantId $TenantId
        } else {
            Connect-MgGraph -Scopes $requiredScopes -UseDeviceCode
        }
    }
}

Ensure-GraphConnection

if (-not $PrincipalObjectId -or -not $PrincipalAppId) {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$PrincipalAppId'" | Select-Object -First 1
    if (-not $sp) {
        throw "Unable to resolve service principal for app id $PrincipalAppId."
    }

    if (-not $PrincipalObjectId) { $PrincipalObjectId = $sp.Id }
    if (-not $PrincipalAppId) { $PrincipalAppId = $sp.AppId }
    if (-not $PrincipalDisplayName) {
        $PrincipalDisplayName = $sp.DisplayName
    }
}

if (-not $PrincipalDisplayName) {
    $PrincipalDisplayName = "Custom Graph App"
}

function Resolve-SiteId {
    param(
        [string]$Name
    )

    $sites = Get-MgSite -Search $Name -Top 5 -Property id,displayName,name,webUrl
    if (-not $sites) {
        throw "No sites matched '$Name'."
    }

    $matches = $sites | Where-Object { $_.DisplayName -eq $Name -or $_.Name -eq $Name }
    if (-not $matches) {
        Write-Warning "No exact match for '$Name'. Using first result."
        $matches = @($sites[0])
    }

    return $matches[0]
}

function Get-FriendlySiteLabel {
    param([Parameter(Mandatory = $true)]$Site)

    if (-not $Site) {
        return "(unknown)"
    }

    $label = $Site.DisplayName
    if ([string]::IsNullOrWhiteSpace($label) -and $Site.AdditionalProperties -and $Site.AdditionalProperties.ContainsKey("displayName")) {
        $label = $Site.AdditionalProperties["displayName"]
    }

    if ([string]::IsNullOrWhiteSpace($label)) {
        $label = $Site.Name
    }
    if ([string]::IsNullOrWhiteSpace($label) -and $Site.AdditionalProperties -and $Site.AdditionalProperties.ContainsKey("name")) {
        $label = $Site.AdditionalProperties["name"]
    }
    if ([string]::IsNullOrWhiteSpace($label)) {
        $label = $Site.Id
    }
    return $label
}

if ($PSCmdlet.ParameterSetName -eq "ByName") {
    $site = Resolve-SiteId -Name $SiteName
    if (-not $site) {
        throw "Failed to resolve site for name '$SiteName'."
    }

    $SiteId = $site.Id
    $friendlyName = Get-FriendlySiteLabel -Site $site
    Write-Host "Resolved search '$SiteName' to site '$friendlyName' (id $SiteId)" -ForegroundColor Cyan
}
elseif ($PSCmdlet.ParameterSetName -eq "ByUrl") {
    try {
        $uri = [Uri]$SiteUrl
    } catch {
        throw "Invalid SiteUrl '$SiteUrl'. Supply a full SharePoint URL."
    }

    $path = $uri.AbsolutePath.TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($path) -or $path -eq "/") {
        throw "SiteUrl must include the site path, e.g. https://contoso.sharepoint.com/sites/testsite"
    }
    if (-not $path.StartsWith("/")) {
        $path = "/$path"
    }

    $graphSiteId = "$($uri.Host):$path"
    try {
        $site = Get-MgSite -SiteId $graphSiteId -Property id,displayName,name,webUrl
    } catch {
        throw "Unable to resolve site from url '$SiteUrl': $($_.Exception.Message)"
    }

    if (-not $site) {
        throw "Graph did not return a site for url '$SiteUrl'."
    }

    $SiteId = $site.Id
    $friendlyName = Get-FriendlySiteLabel -Site $site
    Write-Host "Resolved url '$SiteUrl' to site '$friendlyName' (id $SiteId)" -ForegroundColor Cyan
}
else {
    Write-Host "Using provided site id $SiteId" -ForegroundColor Cyan
}

$payload = @{
    roles = @($Role)
    grantedToIdentities = @(
        @{ application = @{ id = $PrincipalAppId; displayName = $PrincipalDisplayName } }
    )
} | ConvertTo-Json -Depth 5

$permissionUri = "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions"
Write-Host "Granting $Role access to site $SiteId for principal appId $PrincipalAppId (objectId $PrincipalObjectId)..." -ForegroundColor Yellow

try {
    $result = Invoke-MgGraphRequest -Uri $permissionUri -Method POST -Body $payload -ContentType "application/json"
    Write-Host "Grant successful. Permission id: $($result.id)" -ForegroundColor Green
} catch {
    throw "Failed to grant access: $($_.Exception.Message)"
}
