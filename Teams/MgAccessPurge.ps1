#Requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Teams,Microsoft.Graph.Users

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

$scopes = 'User.Read.All', 'TeamSettings.ReadWrite.All', 'ChannelSettings.ReadWrite.All', 'ChannelMember.ReadWrite.All', 'GroupMember.ReadWrite.All', 'TeamMember.ReadWrite.All', 'Domain.Read.All'

$MissingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($MissingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection if asked." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

"INFO`tChecking given users" | Write-Log

$userInfo = Get-MgUserInfo -User $User

if($userInfo){
    "INFO`tUser found: {0}`n`t- ID:`t{1}`n`t- Mail:`t{2}" -f $userInfo.DisplayName,$userInfo.Id, $userInfo.Mail | Write-Log
}else{
    "ERROR`tFailed to find corresponding user."
    return
}

$teams = Get-MgUserJoinedTeam -UserId $userInfo.Id

$i = 0
$iMax = $teams.count

foreach($team in $teams){
    Write-Progress -Activity "Removing access" -Status $team.DisplayName -Id 0 -PercentComplete (($i/$iMax)*100)
    $i++

    "INFO`tStarting channel verification for Team: {0}" -f $team.DisplayName | Write-Log

    $channels = Get-MgTeamChannel -TeamId $team.Id -Filter "membershipType eq 'private'"

    $j = 0
    $jMax = $channels.Count

    $fails = $false

    foreach($channel in $channels){
        if($jMax -gt 3){
            Write-Progress -Activity "Removing channel access" -Status $channel.DisplayName -Id 1 -PercentComplete (($j/$jMax)*100) -ParentId 0
        }
        $j++

        $userChannelInfo = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members?`$filter=displayName eq '$($userInfo.DisplayName)'")['value']
        if($userChannelInfo){
            try{
                # Remove-MgTeamChannelMember -TeamId $team.Id -ChannelId $channel.Id -ConversationMemberId $userChannelInfo.Id -ErrorAction Stop
                Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members/$($userChannelInfo.id)" -ErrorAction Stop
                "INFO`tRemoved user from channel: {0}" -f $channel.DisplayName | Write-Log
            }catch{
                "ERROR`tFailed to remove user from channel: {0}" -f $channel.DisplayName | Write-Log
                $fails = $true
            }
        }
    }
    Write-Progress -Activity "Removing channel access" -Completed

    if($fails){
        "ERROR`tChannel removal failed. Cancelled Team removal." | Write-Log
    }else{
        "INFO`tRemoved from Team: {0}" -f $team.DisplayName | Write-Log
    }
}

"====Script Stop====" | Write-Log


<#
    .SYNOPSIS
    Removes the given users from all Teams and channels he is a member of.

    .DESCRIPTION
    Lists the Teams the user is member of. Goes over the private channels of those Teams and
    removes those user from those before removing the Teams access. If the removal from the
    channel fails, the removal from the parent Team will be skipped.
    All data is logged within a log file in the same library as the script.

    .NOTES
    This script was written due to some users still existing as "ghosts" within channels when
    being directly deleted from Entra. It won't have a reason for existing if that issue is
    solved.

    .PARAMETER User
    This can be the mail, SamAccountName or ObjectId of the targeted user. The given information
    is first matched against the EntraId. If no corresponding user is found the script stops.

    .LINK
    Get-MgUserJoinedTeam

    .LINK
    Invoke-MgGraphRequest

    .EXAMPLE
    PS> .\MgAccessPurge.ps1 -User john.doe@contoso.com

    .EXAMPLE
    PS> .\MgAccessPurge.ps1 -User ff34e1d4-45a7-487c-b95a-f35d21796034
#>

