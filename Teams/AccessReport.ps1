param(
    [Parameter(Mandatory)]
    [String]$User,
    [String]$Path
)
#Requires -Modules MicrosoftTeams
function Show-Info {param ([Parameter(Mandatory,Position=0)][String]$Message)Write-Host "`n(i)$Message" -ForegroundColor Gray}

try{Get-CsTenant|Out-Null}catch{throw "Please connect the Microsoft Teams Services"}

try{
    Show-Info "Verifying user"
    $Teams = Get-Team -User $User -ErrorAction Stop
}catch{
    throw "`'$User`' not found"
}

Show-Info "Generating report"

class TeamsPermission{
    [String]$GroupId
    [String]$Teams
    [String]$Channel
    [String]$Type
    [String]$Access
    TeamsPermission(
        [String]$g,
        [String]$t,
        [String]$c,
        [String]$ty,
        [String]$a
    ){
        $this.GroupId = $g
        $this.Teams = $t
        $this.Channel = $c
        $this.Type = $ty
        $this.Access = $a
    }
}

$i = 0
$iMax = $Teams.Count

$Report = foreach($Team in $Teams){
    Write-Progress -Activity "Documenting Teams[$i/$iMax]" -Status $Team.DisplayName -PercentComplete ($i/$iMax*100) -Id 0
    Write-Host "Teams: $($Team.DisplayName)"
    $Channels = Get-TeamChannel -GroupId $Team.GroupId
    $j = 0
    $jMax = $Channels.Count
    foreach($Channel in $Channels){
        [String]$ChannelName = $Channel.DisplayName
        Write-Progress -Activity "Documenting Channels[$j/$jMax]" -Status $ChannelName -PercentComplete ($j/$jMax*100) -Id 1 -ParentId 0
        $ChannelUsers = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $ChannelName
        if($ChannelUsers.user -contains $User){
            $Role = $ChannelUsers.Role[$ChannelUsers.User.IndexOf($User)]
            [TeamsPermission]::new(
                $Team.GroupId,
                $Team.DisplayName,
                $Channel.DisplayName,
                $Channel.MembershipType,
                $Role
            )
            Write-Host "`tChannel: $($Channel.DisplayName) [$Role]"
        }
        $j++
    }
    Write-Progress -Activity "Documenting Channels" -Id 1 -ParentId 0 -Completed
    $i++
}
Write-Progress -Activity "Documenting Teams" -Completed -Id 0
if(!$Path){
    $Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_$($User.Replace("@","_").Replace(".","_")).csv" 
}elseif(Test-Path -Path $Path -PathType Container){
    $Path += "\$(Get-Date -Format yyyyMMdd)_$($User.Replace("@","_").Replace(".","_")).csv"
    $Path = $Path.Replace("\\","\")
}
$Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
Show-Info "File saved under: $Path"

<#
    .SYNOPSIS

    Documents permissions of a user.

    .DESCRIPTION

    Returns a CSV containing the Teams, Channels and the accesslevel the user has on those.

    .PARAMETER User
    UserPrincipalName of the user you want to know the permissions of.

    .PARAMETER Path
    Filepath for the report.

    .INPUTS

    None. You cannot pipe objects into AccessReports.ps1 .

    .OUTPUTS

    None.

    .EXAMPLE

    PS> AccessReport.ps1 -User 'john.doe@contosco.com'

    .EXAMPLE

    PS> AccessReport.ps1 -User 'john.doe@contosco.com' -Path '.\Report.csv'
    
    .LINK

    Get-Team

    .LINK

    Get-TeamChannelUser

    .LINK

    Get-TeamChannel
#>