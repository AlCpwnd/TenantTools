#Requires -Modules MicrosoftTeams

param(
    [Parameter(Mandatory)]
    [String]$User,
    [String]$Path
)
function Show-Info {param ([Parameter(Mandatory,Position=0)][String]$Message)Write-Host "`t(i)$Message" -ForegroundColor Gray}
function Show-Error {param ([Parameter(Mandatory,Position=0)][String]$Message)Write-Host "`t[x]$Message" -ForegroundColor Red}

try{Get-CsTenant|Out-Null}catch{"";Show-Error "Please connect the Microsoft Teams Services";"";return}

try{
    Show-Info "Verifying user"
    $Teams = Get-Team -User $User -ErrorAction Stop
}catch{
    Show-Error "`'$User`' not found"
    return
}

Show-Info "Generating report"

$i = 0
$iMax = $Teams.Count

$Report = foreach($Team in $Teams){
    Write-Progress -Activity "Documenting Teams" -Status $Team.DisplayName -PercentComplete ($i/$iMax*100) -Id 0
    $Channels = Get-TeamChannel -GroupId $Team.GroupId
    $j = 0
    $jMax = $Channels.Count
    foreach($Channel in $Channels){
        Write-Progress -Activity "Documenting Channels" -Status $Channel.DisplayName -PercentComplete ($j/$jMax*100) -Id 1 -ParentId 0
        $ChannelUsers = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $Channel.DisplayName
        if($ChannelUsers.user -contains $User){
            [PSCustomObject]@{
                GroupId = $Team.GroupId
                Teams = $Team.DisplayName
                Channel = $Channel.DisplayName
                Access = $ChannelUsers.Role[$ChannelUsers.User.IndexOf($User)]
            }
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