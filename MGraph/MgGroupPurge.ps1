#Requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Users,Microsoft.Graph.Groups

param(
    [Parameter(Mandatory = $true)]
    [string]$User
)

function Write-Log {
    param(
        [Parameter(ValueFromPipeLine = $true, Mandatory = $true)]
        # Message to append to the log file
        [System.String]$Text,

        [Parameter()]
        # Path to the log file to write to
        [String]$LogFile = $Global:LogPath
    )
    begin {
        if (-not $LogFile) {
            $LogFile = Split-Path -Leaf $PSCommandPath.Replace('.ps1', '.log')
            $Global:LogPath = $LogFile
        }
        if (-not (Test-Path -Path $logFile)) {
            New-Item -Path $LogFile
            Write-Host "Log file generated: $logFile"
            "Log file generated: $logFile" | Out-File -FilePath $LogFile -Encoding utf8
            "_Structure_________________________", 
            "yyyy-MM-dd - HH:mm:ss`tType`tDetail", 
            "___________________________________" | Out-File -FilePath $LogFile -Encoding utf8 -Append
        }
    }
    process {
        $date = Get-Date -Format "yyyy-MM-dd - HH:mm:ss"
        "$date`t$text" | Out-File -FilePath $logFile -Append -Encoding utf8
    }
}

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

"====Script Start====" | Write-Log

"INFO`tSetting up environment" | Write-Log

$scopes = 'User.Read.All', 'GroupMember.ReadWrite.All', 'Domain.Read.All'

$MissingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($MissingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection if asked." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

"INFO`tChecking given users" | Write-Log

$userInfo = Get-MgUserInfo -User $User

if ($userInfo) {
    "INFO`tUser found:`n`t- Name:`t{0}`n`t- ID:`t{1}`n`t- Mail:`t{2}" -f $userInfo.DisplayName, $userInfo.Id, $userInfo.Mail | Write-Log
}
else {
    "ERROR`tFailed to find corresponding user." | Write-Log
    return
}

# Recovers user groups, but removes Dynamic groups.
$userGroups = Get-MgUserMemberOf -UserId $userInfo.Id | Where-Object { $_.AdditionalProperties['groupTypes'] -notcontains 'DynamicMembership' -and -not $_.AdditionalProperties['mailEnabled'] }

$i = 0
$iMax = $userGroups.count

foreach ($group in $userGroups) {
    Write-Progress -Activity "Removing membership" -Status $group.AdditionalProperties['displayName'] -PercentComplete (($i / $iMax) * 100)
    $i++

    try {
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $userInfo.Id -ErrorAction Stop
        "INFO`tRemoved from: {0}" -f $group.AdditionalProperties['displayName'] | Write-Log
    }
    catch {
        "ERROR`tFailed to remove from: {0}" -f $group.AdditionalProperties['displayName'] | Write-Log
    }
}

"====Script Stop====" | Write-Log


<#
    .SYNOPSIS
    Removes a given user from all the groups he is part of.
    
    .DESCRIPTION
    The script will list the groups the given user is part of and remove him from them with
    exception of:
    - Mail enabled security groups
    - Dynamic membership groups
    All changes will be logged into a file with the same name as the script.

    .NOTES
    This script is meant to be ran after the Teams access purge script.

    .PARAMETER User
    This can be the mail, SamAccountName or ObjectId of the targeted user. The given information
    is first matched against the EntraId. If no corresponding user is found the script stops.

    .LINK
    Get-MgUserMemberOf

    .EXAMPLE
    PS> .\MgGroupPurge -User john.doe@contoso.com
#>

