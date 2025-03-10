#Requires -modules Microsoft.Graph.Teams,Microsoft.Graph.Users

param(
    [Parameter(Mandatory)][String]$User,
    [String]$Path
)

$scopes = 'User.Read.All','TeamSettings.Read.All','ChannelSettings.Read.All','ChannelMember.Read.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "Scopes missing:$($scopes -join ',')"
    Connect-MgGraph -Scopes $MissingScopes
}

class TeamsPermission{
    [String]$GroupId
    [String]$Teams
    [String]$ChannelId
    [String]$Channel
    [String]$Type
    [String]$Access
    TeamsPermission(
        [String]$g,
        [String]$t,
        [String]$ci,
        [String]$c,
        [String]$ty,
        [String]$a
    ){
        $this.GroupId = $g
        $this.Teams = $t
        $this.ChannelId = $ci
        $this.Channel = $c
        $this.Type = $ty
        $this.Access = $a
    }
}

$UserInfo = Get-MgUser -UserId $User
$UserGroups = Get-MgUserMemberOf -UserId $UserInfo.Id

$Teams = Get-MgTeam -All:$true
$UserTeams = $Teams | Where-Object{$UserGroups.Id -contains $_.Id}

$i = 0
$iMax = $UserTeams.Count

$Report = foreach($Team in $UserTeams){
    $i ++
    Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -PercentComplete (($i/$iMax)*100)
    $Channels = Get-MgTeamChannel -TeamId $Team.Id
    $j = 0
    $jMax = $Channels.Count
    foreach($Channel in $Channels){
        $j ++
        Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -PercentComplete (($j/$jMax)*100) -ParentId 0
        $Members = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $Channel.Id
        if($Members.AdditionalProperties.Values -contains $UserInfo.Id){
            $UserPermission = $Members[$Members.DisplayName.IndexOf($UserInfo.DisplayName)].Roles
            if($UserPermission.Roles){
                $Role = $UserPermission.Roles[0]
            }else{
                $Role = 'Member'
            }
            [TeamsPermission]::new(
                $Team.Id,
                $Team.DisplayName,
                $Channel.Id,
                $Channel.DisplayName,
                $Channel.MembershipType,
                $Role
            )
        }
    }
    Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -ParentId 0 -Completed
}
Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -Completed

if(!$Path){
    $Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_$($User.Replace("@","_").Replace(".","_")).csv" 
    $i = 1
    while(Test-Path -Path $Path){
        $Path = $Path.Replace(".csv"," ($i).csv")
        $i ++
    }
}elseif(Test-Path -Path $Path -PathType Container){
    $Path += "\$(Get-Date -Format yyyyMMdd)_$($User.Replace("@","_").Replace(".","_")).csv"
    $Path = $Path.Replace("\\","\")
    $i = 1
    while(Test-Path -Path $Path){
        $Path = $Path.Replace(".csv"," ($i).csv")
        $i ++
    }
}
$Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
Write-Host "File saved under: $Path"