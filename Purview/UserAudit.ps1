#Requires -Modules Microsoft.Graph.Authentication,Microsoft.Graph.Users

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$User
)

function Get-MgUserInfo {
    param(
        [Parameter(Mandatory = $true)]
        # String you want to identify the user by.
        [System.String]$User
    )
    $fields = 'DisplayName', 'ID', 'Mail', 'UserPrincipalName', 'UserType'
    if ($User -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$') {
        $output = Get-MgUser -UserId $User -Property $fields | Select-Object $fields
    }
    if ($User -like "*@*") {
        $domains = (Get-MgDomain).Id -join '|'
        if ($User -match $domains) {
            $output = Get-MgUser -UserId $User -Property $fields | Select-Object $fields
        }
        else {
            $output = Get-MgUser -Search "Mail:$User" -Property $fields -ConsistencyLevel eventual | Select-Object $fields
        }
    }
    return $output
}

$scopes = 'User.Read.All', 'AuditLogsQuery.Read.All', 'Domain.Read.All'

$MissingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($MissingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection if asked." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

"INFO`tChecking given users" | Write-Host

$userInfo = Get-MgUserInfo -User $User

if ($userInfo) {
    "INFO`tUser found:`n`t- Name:`t{0}`n`t- ID:`t{1}`n`t- Mail:`t{2}" -f $userInfo.DisplayName, $userInfo.Id, $userInfo.Mail | Write-Host
}
else {
    throw "Failed to find corresponding user."
}

$startDate = (Get-Date).AddDays(-7).ToUniversalTime()
$endDate   = (Get-Date).ToUniversalTime()

$body = @{
    "@odata.type"       = "#microsoft.graph.security.auditLogQuery"
    displayName         = "{0}-{1}-{2}" -f $startDate.ToString('dd/MM/yyyy'),$endDate.ToString('dd/MM/yyyy'), $userInfo.DisplayName
    filterStartDateTime = $startDate.ToString("o")
    filterEndDateTime   = $endDate.ToString("o")
    
    userPrincipalNameFilters = @(
        $userInfo.UserPrincipalName
    )

    recordTypeFilters = @(
        "sharePointFileOperation"
    )

    operationFilters = @(
        "FileDownloaded"
        "FileDeleted"
        "FileRecycled"
        "FolderRecycled"
        "FileVersionRecycled"
        "FileVersionsAllRecycled"
    )
} | ConvertTo-Json -Depth 10

$auditSearch = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/security/auditLog/queries" -Body $body -ContentType "application/json"

return $auditSearch


<#
    .SYNOPSIS
    Runs an audit of most disruptive SharePoint activities of the given user.
    
    .DESCRIPTION
    Prepares a Purview SharePoint activity audit for the given user over the last 7 days.
    The activities that are being monitored are:
    - File Download
    - File Deletion
    - File Deletion (moved to recycle bin)
    - Folder Deletion (moved to recycle bin)
    - File Version Removal
    The final result of the search is then available at: https://purview.microsoft.com/audit/

    .PARAMETER User
    This can be the mail, SamAccountName or ObjectId of the targeted user. The given information
    is first matched against the EntraId. If no corresponding user is found the script stops.

    .LINK
    https://learn.microsoft.com/en-us/purview/audit-solutions-overview

    .EXAMPLE
    PS> .\UserAudit.ps1 -User john.doe@contoso.com
#>

