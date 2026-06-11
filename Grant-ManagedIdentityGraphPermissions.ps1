[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ManagedIdentityObjectId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$AppPermissions,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$PassThru
)

$ErrorActionPreference = "Stop"

$graphAppId = "00000003-0000-0000-c000-000000000000"
$requiredScopes = @(
    "AppRoleAssignment.ReadWrite.All",
    "Application.Read.All"
)

function Connect-ToGraph {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        if ([string]::IsNullOrWhiteSpace($TenantId)) {
            Connect-MgGraph -Scopes $requiredScopes 
        }
        else {
            Connect-MgGraph -TenantId $TenantId -Scopes $requiredScopes 
        }
    }
}

Connect-ToGraph

try {
    $managedIdentitySp = Get-MgServicePrincipal -ServicePrincipalId $ManagedIdentityObjectId
}
catch {
    throw "Managed identity service principal '$ManagedIdentityObjectId' was not found."
}

$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
if (-not $graphSp) {
    throw "Microsoft Graph service principal was not found in this tenant."
}

$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySp.Id -All
$results = @()

foreach ($permission in $AppPermissions) {
    $appRole = $graphSp.AppRoles | Where-Object {
        $_.Value -eq $permission -and
        $_.AllowedMemberTypes -contains "Application" -and
        $_.IsEnabled
    }

    if (-not $appRole) {
        Write-Warning "Permission '$permission' is not a valid Microsoft Graph application permission in this tenant."
        continue
    }

    $alreadyAssigned = $existingAssignments | Where-Object {
        $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $appRole.Id
    }

    if ($alreadyAssigned) {
        Write-Host "Skipping '$permission' (already assigned)."
        $results += [pscustomobject]@{
            Permission = $permission
            Status     = "AlreadyAssigned"
            AppRoleId  = $appRole.Id
        }
        continue
    }

    $body = @{
        AppRoleId   = $appRole.Id
        PrincipalId = $managedIdentitySp.Id
        ResourceId  = $graphSp.Id
    }

    if ($PSCmdlet.ShouldProcess($managedIdentitySp.DisplayName, "Assign Microsoft Graph app role '$permission'")) {
        $assignment = New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $managedIdentitySp.Id -BodyParameter $body
        Write-Host "Assigned '$permission'."

        $results += [pscustomobject]@{
            Permission   = $permission
            Status       = "Assigned"
            AssignmentId = $assignment.Id
            AppRoleId    = $assignment.AppRoleId
        }
    }
}

if ($PassThru) {
    $results
}
else {
    $results | Format-Table -AutoSize
}

<#
Example:

.\Grant-ManagedIdentityGraphPermissions.ps1 \
  -ManagedIdentityObjectId "11111111-2222-3333-4444-555555555555" \
  -AppPermissions "User.Read.All","Group.Read.All" \
  -TenantId "contoso.onmicrosoft.com"
#>