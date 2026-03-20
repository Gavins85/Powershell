<#
.SYNOPSIS
  Remove specified Microsoft Graph application permissions (app roles) from a service principal.

.DESCRIPTION
  Uses Microsoft.Graph PowerShell to:
   - Connect (interactive) if needed
   - Find the target service principal by objectId or displayName
   - Find Microsoft Graph's service principal (appId 00000003-0000-0000-c000-000000000000)
   - Remove app role assignments that match the requested permission values

.NOTES
  - Assumes Microsoft.Graph module is already installed.
  - Requires sufficient admin privileges (e.g. Global Administrator or Privileged Role Admin).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $TargetDisplayName = "", #add name of EnterPrise App

    [Parameter(Mandatory = $false)]
    [string] $TargetObjectId,

    [Parameter(Mandatory = $false)]
    [string[]] $PermissionsToRemove = @(
        "Calendars.Read",
        "Calendars.ReadWrite",
        "IdentityRiskyUser.Read.All"
    ),

    [Parameter(Mandatory = $false)]
    [switch] $WhatIf,

    [Parameter(Mandatory = $false)]
    [switch] $Force
)

try {
    # Import module (assumes installed)
    Import-Module Microsoft.Graph -ErrorAction Stop

    # Connect if no context
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        $scopes = @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All")
        Write-Host "Connecting to Microsoft Graph (interactive)..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes $scopes
    }

    # 1) Locate target service principal (prefer objectId if provided)
    if ($TargetObjectId) {
        $sp = Get-MgServicePrincipal -ServicePrincipalId $TargetObjectId -ErrorAction SilentlyContinue
        if (-not $sp) { Throw "No service principal found with objectId $TargetObjectId" }
    }
    else {
        $results = Get-MgServicePrincipal -Filter "displayName eq '$($TargetDisplayName)'" -ConsistencyLevel eventual -ErrorAction Stop
        if (-not $results) { Throw "No service principal found with displayName '$TargetDisplayName'." }
        if ($results.Count -gt 1) {
            Write-Warning "Multiple service principals matched '$TargetDisplayName'. Using the first result. Consider supplying -TargetObjectId for exact match."
        }
        $sp = $results[0]
    }
    Write-Host "Target service principal: $($sp.DisplayName) (Id: $($sp.Id))" -ForegroundColor Green

    # 2) Microsoft Graph service principal
    $graphAppId = "00000003-0000-0000-c000-000000000000"
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -ErrorAction Stop
    if (-not $graphSp) { Throw "Microsoft Graph service principal not found (appId $graphAppId)." }
    Write-Host "Microsoft Graph SP Id: $($graphSp.Id)" -ForegroundColor Green

    # 3) Map permission values -> AppRole Ids
    $targetAppRoles = @{}
    foreach ($role in $graphSp.AppRoles) {
        if ($role.Value -and ($PermissionsToRemove -contains $role.Value)) {
            # AppRoles Id may be a Guid or string - keep as string
            $targetAppRoles[$role.Id.ToString()] = $role.Value
        }
    }

    if ($targetAppRoles.Count -eq 0) {
        Write-Warning "None of the specified permission values were found on Microsoft Graph's app roles. Exiting."
        return
    }

    Write-Host "App roles to remove:" -ForegroundColor Cyan
    $targetAppRoles.GetEnumerator() | ForEach-Object { Write-Host " - $($_.Value) -> AppRoleId: $($_.Key)" }

    # 4) Get current app role assignments for the target SP pointing to Microsoft Graph
    $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction Stop |
    Where-Object { $_.ResourceId -eq $graphSp.Id }

    if (-not $assignments) {
        Write-Host "No app role assignments to Microsoft Graph found for this service principal." -ForegroundColor Yellow
        return
    }

    # 5) Filter assignments to those matching our roles
    $matches = @()
    foreach ($a in $assignments) {
        $appRoleIdStr = $a.AppRoleId.ToString()
        if ($targetAppRoles.ContainsKey($appRoleIdStr)) {
            $matches += [PSCustomObject]@{
                AssignmentId   = $a.Id
                AppRoleId      = $appRoleIdStr
                PermissionName = $targetAppRoles[$appRoleIdStr]
                PrincipalId    = $a.PrincipalId
            }
        }
    }

    if ($matches.Count -eq 0) {
        Write-Host "No matching app role assignments found for the specified permissions." -ForegroundColor Yellow
        return
    }

    Write-Host "`nMatching assignments found:" -ForegroundColor Cyan
    $matches | Format-Table -AutoSize

    if ($WhatIf) {
        Write-Host "`nWhatIf: no changes will be made. The script would remove the above assignments." -ForegroundColor Yellow
        return
    }

    if (-not $Force) {
        $confirm = Read-Host "Proceed to remove the above $($matches.Count) assignment(s)? Type 'YES' to continue"
        if ($confirm -ne "YES") {
            Write-Host "Aborting - user did not confirm." -ForegroundColor Yellow
            return
        }
    }

    # 6) Remove assignments
    $errors = @()
    foreach ($m in $matches) {
        try {
            Write-Host "Removing assignment $($m.AssignmentId) ($($m.PermissionName))..." -ForegroundColor Gray
            Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -AppRoleAssignmentId $m.AssignmentId -ErrorAction Stop
            Write-Host "Removed: $($m.PermissionName) ($($m.AssignmentId))" -ForegroundColor Green
        }
        catch {
            Write-Host "ERROR removing $($m.PermissionName): $($_.Exception.Message)" -ForegroundColor Red
            $errors += $_
        }
    }

    # 7) Verification
    $remaining = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction Stop |
    Where-Object { $_.ResourceId -eq $graphSp.Id }

    Write-Host "`nRemaining Microsoft Graph app role assignments for this SP:" -ForegroundColor Cyan
    if ($remaining) {
        $remaining | ForEach-Object {
            $rName = ($graphSp.AppRoles | Where-Object { $_.Id.ToString() -eq $_.AppRoleId.ToString() }).Value
            if (-not $rName) { $rName = $_.AppRoleId }
            Write-Host " - AssignmentId: $($_.Id)  AppRoleId: $($_.AppRoleId)  Name: $rName"
        }
    }
    else {
        Write-Host "None." -ForegroundColor Green
    }

    if ($errors.Count -gt 0) {
        Write-Host "`nCompleted with errors. Inspect messages above." -ForegroundColor Red
        exit 1
    }
    else {
        Write-Host "`nCompleted successfully." -ForegroundColor Green
    }
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
    throw
}