
# Object ID of the Enterprise App (service principal) to grant permissions to
$ObjectId = "e70f622c-57a7-45dc-938f-6378c9fde21f"
# Application Permissions (app roles) to grant, space-separated
$graphScope = "Application.Read.All Directory.Read.All"

Connect-MgGraph -Scope AppRoleAssignment.ReadWrite.All

# Get the Microsoft Graph service principal (resource)
$graph = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"

# Grant each Application Permission (app role) at the tenant level
foreach ($scope in ($graphScope -split " ")) {
    $appRole = $graph.AppRoles | Where-Object { $_.Value -eq $scope -and $_.AllowedMemberTypes -contains "Application" }

    if (-not $appRole) {
        Write-Warning "App role '$scope' not found in Microsoft Graph."
        continue
    }

    $params = @{
        "AppRoleId"   = $appRole.Id
        "PrincipalId" = $ObjectId
        "ResourceId"  = $graph.Id
    }

    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ObjectId -BodyParameter $params |
    Format-List Id, AppRoleId, PrincipalId, ResourceId, CreatedDateTime
}
