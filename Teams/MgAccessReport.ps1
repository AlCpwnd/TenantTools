#Requires -modules Microsoft.Graph.Teams,Microsoft.Graph.Users

param(
    [Parameter(Mandatory,ParameterSetName='Single')][String]$User,
    [Parameter(Mandatory,ParameterSetName='All')][Switch]$All,
    [String]$Path
)

$scopes = 'User.Read.All','TeamSettings.Read.All','ChannelSettings.Read.All','ChannelMember.Read.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object{Write-Host "`t> $_" -ForegroundColor Yellow}
    Write-Host "Adding scopes to current environment. Please allow the connection." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

class TeamsPermission{
    [String]$GroupId
    [String]$Teams
    [String]$ChannelId
    [String]$Channel
    [String]$Type
    [String]$User
    [String]$Access
    TeamsPermission(
        [String]$g,
        [String]$t,
        [String]$ci,
        [String]$c,
        [String]$ty,
        [string]$u,
        [String]$a
    ){
        $this.GroupId = $g
        $this.Teams = $t
        $this.ChannelId = $ci
        $this.Channel = $c
        $this.Type = $ty
        $this.User = $u
        $this.Access = $a
    }
}

if($PSCmdlet.ParameterSetName -eq "Single"){
    $UserInfo = Get-MgUser -UserId $User
    $Teams = Get-MgUserJoinedTeam -UserId $UserInfo.Id
}else{
    $Teams = Get-MgTeam -All:$true
}

$i = 0
$iMax = $Teams.Count

$Report = foreach($Team in $Teams){
    $i ++
    Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -PercentComplete (($i/$iMax)*100)
    $Channels = Get-MgTeamChannel -TeamId $Team.Id
    $j = 0
    $jMax = $Channels.Count
    foreach($Channel in $Channels){
        $j ++
        # Prevents flickering.
        if($jMax -gt 3){
            Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -PercentComplete (($j/$jMax)*100) -ParentId 0
        }
        # Get-MgChannelMember command isn't working atm. Going through API calls as an alternative.
        # $Members = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $Channel.Id
        $Members = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($Team.Id)/channels/$($Channel.Id)/members").value
        if(
            $Members.AdditionalProperties.Values -contains $UserInfo.Id -and
            $PSCmdlet.ParameterSetName -eq "Single"
        ){
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
                $UserInfo.DisplayName,
                $Role
            )
        }else{
            foreach($member in $Members){
                if($member.Roles){
                    $Role = $member.Roles[0]
                }else{
                    $Role = 'Member'
                }
                [TeamsPermission]::new(
                    $Team.Id,
                    $Team.DisplayName,
                    $Channel.Id,
                    $Channel.DisplayName,
                    $Channel.MembershipType,
                    $member.DisplayName,
                    $Role
                )
            }
        }
    }
    Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -ParentId 0 -Completed
}
Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -Completed

if(!$Path){
    if($PSCmdlet.ParameterSetName -eq "Single"){
        $Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_$($User.Replace("@","_").Replace(".","_")).csv" 
    }else{
        $Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_All.csv"
    }
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
Write-Host "File saved under: $Path" -ForegroundColor Green