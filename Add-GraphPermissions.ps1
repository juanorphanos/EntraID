Connect-MgGraph -Scopes Application.Read.All, AppRoleAssignment.ReadWrite.All, RoleManagement.ReadWrite.Directory, Application.ReadWrite.All, DelegatedPermissionGrant.ReadWrite.All
 
$graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$clientSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph Command Line Tools'"
$principalId = "15f32b50-5efa-4e30-97c7-91932ce961dd" #Object ID del Usuario a agregar
$scope = "Files.ReadWrite.All"
 
$params = @{
    ClientId    = $clientSp.Id    
    ConsentType = 'Principal'
    PrincipalId = $principalId
    ResourceId  = $graphSp.Id      
    Scope       = $scope
}
 

New-MgOauth2PermissionGrant -BodyParameter $params | Format-List Id, ClientId, ConsentType, ExpiryTime, PrincipalId, ResourceId, Scope