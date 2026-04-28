#Requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Teams,Microsoft.Graph.Users

param(
    [Parameter()]
    # DisplayName of GUID of the Teams you want to copy the permissions from.
    # This is used if you want to copy a user's permission on a single Team.
    [System.String]$TargetTeam,

    [Parameter(Mandatory = $true)]
    # Template user of which the permissions need to be copied.
    [System.String]$Template,

    [Parameter(Mandatory = $true)]
    # Target user to which the permissions need to be applied.
    [System.String]$Target,

    [Parameter()]
    # If the target also needs to be added to the private channels.
    [Switch]$CopyChannels,

    [Parameter()]
    # Also copy the user's role within the Teams and channels.
    [Switch]$IncludeRole
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

$templateInfo = Get-MgUserInfo -User $Template
if (-not $templateInfo) {
    Write-Host "Failed to find a user for: $Template" -ForegroundColor Red
    "ERROR`tFailed to find a user for: $Template" | Write-Log
    "====Script Stop====" | Write-Log
    return
}

$targetInfo = Get-MgUserInfo -User $Target
if (-not $targetInfo) {
    Write-Host "Failed to find a user for: $Target" -ForegroundColor Red
    "ERROR`tFailed to find a user for: $Template" | Write-Log
    "====Script Stop====" | Write-Log
    return
}

if ($TargetTeam) {
    if ($TargetTeam -match '^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$') {
        $teamInfo = Get-MgTeam -TeamId $TargetTeam | Select-Object DisplayName, Id
    }
    else {
        $teamInfo = Get-MgTeam -Filter "displayName eq '$TargetTeam'" | Select-Object DisplayName, Id
    }
    if (-not $teamInfo) {
        $errorMessage = "Failed to find '$TargetTeam' in existing Teams. Please confirm this is a valid Team display name or GUID."
        Write-Host $errorMessage -ForegroundColor Red
        "ERROR`t$errorMessage" | Write-Log
        "====Script Stop====" | Write-Log
    }
    "INFO`tCopying existing permissions for Teams: $($TeamInfo.DisplayName)" | Write-Log
    $templateTeams = @($TeamInfo)
}
else {
    "INFO`tRecovering Teams access for template: $($templateInfo.DisplayName)" | Write-Log
    $templateTeams = Get-MgUserJoinedTeam -UserId $templateInfo.Id -All:$true | Select-Object DisplayName, Id
}

"INFO`tRecovering Teams access for target: $($targetInfo.DisplayName)" | Write-Log
$targetTeams = Get-MgUserJoinedTeam -UserId $targetInfo.Id -All:$true | Select-Object DisplayName, Id

"INFO`tStarting addition to Teams" | Write-Log

foreach ($team in $templateTeams) {
    $params = @{
        "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
        roles             = @()
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($targetInfo.Id)')"
    }
    if ($targetInfo.UserType -eq 'Guest') {
        $params.roles += 'guest'
    }
    if ($IncludeRole -and $params.roles -notcontains 'guest') {
        $templateRole = (Get-MgTeamMember -TeamId $team.Id -Filter "(microsoft.graph.aadUserConversationMember/userId eq '$($templateInfo.Id)')").Roles
        if ($templateRole -eq 'owner') {
            $params.roles += 'owner'
        }
        if ($targetTeams.Id -contains $team.Id) {
            $targetTeamInfo = Get-MgTeamMember -TeamId $team.Id -Filter "(microsoft.graph.aadUserConversationMember/userId eq '$($targetInfo.Id)')"
            if ($targetTeamInfo.Roles -notcontains 'owner') {
                "EDIT`tPromoted user to owner for: $($team.DisplayName)" | Write-Log
                $params.Remove('user@odata.bind')
                Update-MgTeamMember -TeamId $team.Id -ConversationMemberId $targetTeamInfo.Id -BodyParameter $params
            }
            else {
                continue
            }
        }
        else {
            "ADD`t`tUser added to: $($team.DisplayName) :as: $(if(-not $params.roles){'member'}else{'owner'})" | Write-Log
            New-MgTeamMember -TeamId $team.Id -BodyParameter $params
        }
    }
    elseif ($targetTeams.Id -notcontains $team.Id) {
        "ADD`t`tUser added to: $($team.DisplayName)" | Write-Log
        New-MgTeamMember -TeamId $team.Id -BodyParameter $params
    }
}

if ($CopyChannels) {
    "INFO`tStarting channel addition" | Write-Log
    foreach ($team in $templateTeams) {
        $channels = Get-MgTeamChannel -TeamId $team.Id -Filter "membershipType eq 'private'" -All
        foreach ($channel in $channels) {
            $params = @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                roles             = @()
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($targetInfo.Id)')"
            }
            if ($targetInfo.UserType -eq 'Guest') {
                $params.roles += 'guest'
            }
            $templateChannelInfo = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members?`$filter=displayName eq '$($templateInfo.DisplayName)'")['value']
            if (-not $templateChannelInfo) {
                # Template user isn't part of the channel.
                continue
            }
            $targetChannelInfo = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members?`$filter=displayName eq '$($targetInfo.DisplayName)'")['value']
            if ($IncludeRole -and $params.roles -notcontains 'guest') {
                if ($templateChannelInfo.Roles -contains 'owner') {
                    $params.roles += 'owner'
                }
                if ($targetChannelInfo) {
                    if ($targetChannelInfo.roles -contains 'owner') {
                        continue
                    }
                    "EDIT`t[$($team.DisplayName)]>$($channel.DisplayName): promoted to owner" | Write-Log
                    $params.Remove('user@odata.bind')
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members/$($targetChannelInfo.Id)" -Body $params
                }
                else {
                    "ADD`t`t[$($team.DisplayName)]>$($channel.DisplayName) :as: $(if(-not $params.roles){'member'}else{'owner'})" | Write-Log
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members" -Body $params
                }
            }
            elseif ($targetChannelInfo) {
                continue
            }
            else {
                "ADD`t`t[$($team.DisplayName)]>$($channel.DisplayName)" | Write-Log
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($channel.Id)/members" -Body $params
            }
        }
    }
}

"====Script Stop====" | Write-Log


<#
    .SYNOPSIS
    Copies the Teams permissions from a template onto a target.

    .DESCRIPTION
    The script recovers the template user's Teams access and applies the same to the target user.
.NOTES
    Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
    Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
    Test-MyTestFunction -Verbose
    Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>

