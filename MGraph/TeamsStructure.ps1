#Requires -modules Microsoft.Graph.Teams

$scopes = 'TeamSettings.Read.All','ChannelSettings.Read.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "Scopes missing:$($scopes -join ',')"
    Connect-MgGraph -Scopes $MissingScopes
}

class TeamsPermission{
    [String]$GroupId
    [String]$Teams
    [String]$TeamType
    [String]$ChannelId
    [String]$Channel
    [String]$ChannelType
    TeamsPermission(
        [String]$g,
        [String]$t,
        [String]$tt,
        [String]$ci,
        [String]$c,
        [String]$ct
    ){
        $this.GroupId = $g
        $this.Teams = $t
        $this.TeamType = $tt
        $this.ChannelId = $ci
        $this.Channel = $c
        $this.ChannelType = $ct
    }
}


$Teams = Get-MgTeam -All:$true

$i = 0
$iMax = $Teams.Count

$Report = foreach($Team in $Teams){
    $i ++
    Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -PercentComplete (($i/$iMax)*100)
    $Channels = Get-MgTeamChannel -TeamId $Team.Id
    $j = 0
    $jMax = $Channels.Count
    foreach($Channel in $Channels){
        if($jMax -gt 5){
            # To prevent the progress bar from flickering.
            $j ++
            Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -PercentComplete (($j/$jMax)*100) -ParentId 0
        }
        [TeamsPermission]::new(
            $Team.Id,
            $Team.DisplayName,
            $Team.Visibility,
            $Channel.Id,
            $Channel.DisplayName,
            $Channel.MembershipType
        )
    }
    Write-Progress -Activity "Documenting Channels [$j/$jMax]" -Status $Channel.DisplayName -Id 1 -ParentId 0 -Completed
}
Write-Progress -Activity "Documenting Teams [$i/$iMax]" -Status $Team.DisplayName -Id 0 -Completed


$Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_TeamsStructure.csv" 
$i = 1
while(Test-Path -Path $Path){
    $Path = $Path.Replace(".csv"," ($i).csv")
    $i ++
}
$Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
Write-Host "File saved under: $Path"

<#
    .SYNOPSIS

    Reports the current Teams and Channels structure.

    .DESCRIPTION

    Returns a CSV containing the Teams, their access type, the Channel 
    and their access type. Using the Microsoft Graph API

    .INPUTS

    None. You cannot pipe objects into TeamsStructure.ps1 .

    .OUTPUTS

    A CSV file containing the requested data.

    .EXAMPLE

    PS> TeamsStructure.ps1

    .LINK

    Get-MgTeam

    .LINK

    Get-MgTeamChannel
#>