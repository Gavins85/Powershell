# Connect with the required delegated permissions
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All"

# Optional: confirm current context
Get-MgContext

# Target service principal object ID (NOT appId/clientId)
$targetSpId = ""

# Example Microsoft Graph application permission name
# Replace with the app role you actually want, e.g. User.Read.All, Group.Read.All, Directory.Read.All
$permissionValue = "User.Read.All"

# Get the Microsoft Graph resource service principal
$graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'"

if (-not $graphSp) {
    throw "Microsoft Graph service principal not found."
}

# Get the client service principal that will receive the app role
$clientSp = Get-MgServicePrincipal -ServicePrincipalId $targetSpId

if (-not $clientSp) {
    throw "Target service principal not found: $targetSpId"
}

# Find the application permission (app role) on Microsoft Graph
# AllowedMemberTypes must include 'Application'
$appRole = $graphSp.AppRoles |
Where-Object {
    $_.Value -eq $permissionValue -and
    $_.AllowedMemberTypes -contains "Application" -and
    $_.IsEnabled -eq $true
} |
Select-Object -First 1

if (-not $appRole) {
    throw "App role '$permissionValue' was not found on Microsoft Graph or is not an application permission."
}

Write-Host "Assigning app role '$($appRole.Value)'"
Write-Host "AppRoleId: $($appRole.Id)"
Write-Host "Client SP: $($clientSp.DisplayName) ($($clientSp.Id))"
Write-Host "Resource SP: $($graphSp.DisplayName) ($($graphSp.Id))"

# Check whether assignment already exists
$existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSp.Id -All |
Where-Object {
    $_.ResourceId -eq $graphSp.Id -and
    $_.AppRoleId -eq $appRole.Id
} |
Select-Object -First 1

if ($existingAssignment) {
    Write-Host "Assignment already exists."
}
else {
    New-MgServicePrincipalAppRoleAssignment `
        -ServicePrincipalId $clientSp.Id `
        -PrincipalId $clientSp.Id `
        -ResourceId $graphSp.Id `
        -AppRoleId $appRole.Id

    Write-Host "Assignment created."
}

# Verify the assignment
Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $clientSp.Id -All |
Where-Object {
    $_.ResourceId -eq $graphSp.Id
} |
Select-Object Id, PrincipalDisplayName, ResourceDisplayName, AppRoleId