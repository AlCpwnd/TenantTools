#Requires -Modules Microsoft.Graph.Groups
[CmdletBinding()]
param(
    [Parameter()][switch]$ExpandMembers
)

$scopes = "RoleManagement.Read.Directory", "Group.Read.All"

$missingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($missingScopes) {
    Write-Host "The following scopes are missing:" -ForegroundColor Yellow
    $missingScopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Connect-MgGraph -Scopes $MissingScopes
}

$adminGroups = Get-MgGroup -Filter "isAssignableToRole eq true" -All:$true
$assignMents = Get-MgRoleManagementDirectoryRoleAssignment -All:$true

$report = foreach ($group in $adminGroups) {
    $roles = $assignMents | Where-Object { $_.PrincipalId -eq $group.Id }
    foreach ($role in $roles) {
        $roleInfo = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $role.RoleDefinitionId
        if ($ExpandMembers) {
            $members = Get-MgGroupMember -GroupId $group.Id
            foreach ($member in $members) {
                [PSCustomObject]@{
                    Group          = $group.DisplayName
                    Role           = $roleInfo.DisplayName
                    RoleDefinition = $roleInfo.Description
                    User           = $member.AdditionalProperties['userPrincipalName']
                }
            }
        }
        else {
            [PSCustomObject]@{
                Group          = $group.DisplayName
                Role           = $roleInfo.DisplayName
                RoleDefinition = $roleInfo.Description
            }
        }
    }
}

if (Get-Module -Name ImportExcel -ListAvailable) {
    $path = $PSCommandPath.Replace('.ps1', "$(Get-Date -Format '_yyyyMMdd_HHmmss').xlsx")
    $report | Export-Excel -Path $Path -WorksheetName Data -TableName Data -TableStyle Medium5 -Show -AutoSize
}
else {
    $path = $PSCommandPath.Replace('.ps1', "$(Get-Date -Format '_yyyyMMdd-HH:mm:ss').csv")
    $report | Export-Csv -Path $path -Encoding utf8 -NoTypeInformation -NoClobber
}
Write-Host "Report saved under: $path" -ForegroundColor green


<#
    .SYNOPSIS
    Returns a report of groups to which admin permissions are assigned to.

    .DESCRIPTION
    Returns an Excel of CSV file with all groups which have admin permissions assigned to them.
    If the "ExpandMembers" switch is chosen, the members of those groups will also be returned.
    
    .NOTES
    Excel version will only be returned if the "ImportExcel" module is detected.
    If the module isn't detected, the report will be written to a CSV file.
    
    .LINK
    Get-MgGroup

    .LINK
    Get-MgRoleManagementDirectoryRoleAssignment

    .LINK
    Get-MgRoleManagementDirectoryRoleDefinition

    .LINK
    Get-MgGroupMember
#>

