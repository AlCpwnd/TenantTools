#Requires -modules Microsoft.Graph.Authentication,Microsoft.Graph.Teams,Microsoft.Graph.Users

param(
    [Parameter(Mandatory=$true)]
    # Template user of which the permissions need to be copied.
    [System.String]$Template,

    [Parameter(Mandatory=$true)]
    # Target user to which the permissions need to be applied.
    [System.String]$Target,

    [Parameter()]
    # If the target also needs to be added to the private channels.
    [Switch]$Channels,

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
        if(-not $LogFile){
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

"====Script Start====" | Write-Log

"INFO`tSetting up environment" | Write-Log

$scopes = 'User.Read.All', 'TeamSettings.ReadWrite.All', 'ChannelSettings.ReadWrite.All', 'ChannelMember.ReadWrite.All', 'GroupMember.ReadWrite.All', 'TeamMember.ReadWrite.All'

$MissingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($MissingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection if asked." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

"INFO`tChecking given users" | Write-Log

try{
    $templateInfo = Get-MgUser -UserId $Template -ErrorAction Stop
}catch{
    Write-Host "Failed to find a user for: $Template" -ForegroundColor Red
    "ERROR`tFailed to find a user for: $Template" | Write-Log
    "====Script Stop====" | Write-Log
    return
}

try{
    $targetInfo = Get-MgUser -UserId $Target -ErrorAction Stop
}catch{
    Write-Host "Failed to find a user for: $Target" -ForegroundColor Red
    "ERROR`tFailed to find a user for: $Template" | Write-Log
    "====Script Stop====" | Write-Log
    return
}

"INFO`tRecovering Teams access for template: $($templateInfo.DisplayName)" | Write-Log
$templateTeams = Get-MgUserJoinedTeam -UserId $templateInfo.Id -All:$true

"INFO`tRecovering Teams access for target: $($targetInfo.DisplayName)" | Write-Log
$targetTeams = Get-MgUserJoinedTeam -UserId $targetInfo.Id -All:$true

"INFO`tStarting addition to Teams" | Write-Log

foreach($team in $templateTeams){
    $params = @{
        "@odata.type" = "#microsoft.graph.aadUserConversationMember"
        roles = @()
        "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($targetInfo.Id)')"
    }
    if($IncludeRole){
        $templateRole = (Get-MgTeamMember -TeamId $team.Id -Filter "(microsoft.graph.aadUserConversationMember/userId eq '$($templateInfo.Id)')").Roles
        if($templateRole -eq 'owner'){
            $params.roles += 'owner'
        }
        if($targetTeams.Id -contains $team.Id){
            $targetTeamInfo = Get-MgTeamMember -TeamId $team.Id -Filter "(microsoft.graph.aadUserConversationMember/userId eq '$($targetInfo.Id)')"
            if($targetTeamInfo.Roles -notcontains 'owner'){
                "EDIT`tPromoted user to owner for: $($team.DisplayName)" | Write-Log
                $params.Remove('user@odata.bind')
                Update-MgTeamMember -TeamId $team.Id -ConversationMemberId $targetTeamInfo.Id -BodyParameter $params
            }else{
                continue
            }
        }
        else{
            "ADD`tUser added to: $($team.DisplayName) :as: $(if(-not $params.roles){'member'}else{'owner'})" | Write-Log
            New-MgTeamMember -TeamId $team.Id -BodyParameter $params
        }
    }elseif($targetTeams.Id -notcontains $team.Id){
        "ADD`tUser added to: $($team.DisplayName)" | Write-Log
        New-MgTeamMember -TeamId $team.Id -BodyParameter $params
    }
}

if($Channels){
    "INFO`tStarting channel addition" | Write-Log
    foreach($team in $targetTeams){
        $channels = Get-MgTeamChannel -TeamId $team.Id -Filter "membershipType eq 'private'" -All:$true
        foreach($Channel in $Channels){
            $params = @{
                "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                roles = @()
                "user@odata.bind" = "https://graph.microsoft.com/v1.0/users('$($targetInfo.Id)')"
            }
            $uri = "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($Channel.Id)/members?`$filter=displayName eq '$($targetInfo.DisplayName)' or displayName eq '$($templateInfo.DisplayName)'"
            $members = (Invoke-MgGraphRequest -Method GET -Uri $uri)['value']
            if(-not $members){
                # Template and target aren't part of the channel.
                continue
            }
            $templateChannelInfo = $members | Where-Object{$_.DisplayName -eq $templateInfo.DisplayName}
            if(-not $templateChannelInfo){
                # Template user isn't part of the channel.
                continue
            }
            if($IncludeRole){                
                if($templateChannelInfo.Roles -contains 'owner'){
                    $params.roles += 'owner'
                }
                if($members.Count -eq 2){
                    $targetChannelInfo = $members | Where-Object{$_.DisplayName -eq $targetInfo.DisplayName}
                    if($targetChannelInfo.roles -contains 'owner'){
                        continue
                    }
                    "EDIT`t[$($team.DisplayName)]>$($Channel.DisplayName): promoted to owner" | Write-Log
                    $params.Remove('user@odata.bind')
                    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($Channel.Id)/members/$($targetChannelInfo.Id)" -Body $params
                }else{
                    "ADD`t[$($team.DisplayName)]>$($Channel.DisplayName) :as: $(if(-not $params.roles){'member'}else{'owner'})" | Write-Log
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($Channel.Id)/members" -Body $params
                }
            }else{
                "ADD`t[$($team.DisplayName)]>$($Channel.DisplayName)" | Write-Log
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$($team.Id)/channels/$($Channel.Id)/members" -Body $params
            }
        }
    }
}
