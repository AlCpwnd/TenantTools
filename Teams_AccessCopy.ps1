#Requires -Modules MicrosoftTeams

param(
    [Parameter(Mandatory = $true)]
    [String]$TemplateUser,
    [Parameter(Mandatory = $true)]
    [string]$Identity,
    [Switch]$IncludeRights,
    [Switch]$Select
)

# Fuctions : ##############################################

function Get-TeamsChannelAccess {
    param(
        [Parameter(Mandatory = $true)]
        [String]$User
    )

    $Teams = Get-Team -User $User

    $i = 0
    $iMax = $Teams.count

    $Report = foreach($Team in $Teams){
        Write-Progress -Activity "Teams :" -Status $Team.DisplayName -Id 0 -PercentComplete (($i/$iMax)*100)
        $Channels = Get-TeamChannel -GroupId $Team.GroupId
        $j = 0
        $jMax = $Channels.Count
        foreach($Channel in $Channels){
            Write-Progress -Activity "Channel :" -Status $Channel.DisplayName -Id 1 -PercentComplete (($j/$jMax)*100) -ParentId 0
            $ChannelPermissions = Get-TeamChannelUser -GroupId $Team.GroupId -DisplayName $Channel.DisplayName
            if($ChannelPermissions.User -contains $User){
                $UserInfo = $ChannelPermissions[$ChannelPermissions.User.IndexOf($User)]
                [PsCustomObject]@{
                    GroupId = $Team.GroupId
                    Teams = $Team.DisplayName
                    Channel = $Channel.DisplayName
                    Type = $Channel.MembershipType
                    Access = $UserInfo.Role
                }
            }
            $j++
        }
        $i++
    }

    return $Report

    <#
        .SYNOPSIS
        Lists the given user's Teams Channel access.

        .DESCRIPTION
        Return a report containing the given user's access to the channels contained within
        those Teams.

        .PARAMETER Teams
        Array containing the Teams access of the user. (Resulting of the Get-Team command)

        .PARAMETER User
        Email address of the user for which the report should be generated

        .INPUTS
        None. You cannot pipe objects to Get-TeamChannelAccess.

        .OUTPUTS
        System.Array. Add-Extension returns an array detailing the user's channel access.

        .EXAMPLE
        PS> Get-TeamsChannelAccess -Teams $Teams -User John.doe@contosco.com

        .LINK
        Get-TeamChannel

        .LINK
        Get-TeamChannelUser
    #>
}

###########################################################


# Input Verification : ####################################

# Lists the $Identity user's existing Teams access.
try{
    Write-Host "`n`t(i):Verifying [$Identity] current access." -ForegroundColor Cyan
    Get-Team -User $Identity -ErrorAction Stop
}
catch{
    Write-Host "`t[x]:$Identity couldn't be found." -ForegroundColor Red
    return
}

# Lists the $TemplateUser user's existing Teams access.
try{
    Write-Host "`n`t(i):Recovering [$TemplateUser] current access." -ForegroundColor Cyan
    $TemplateTeams = Get-TeamsChannelAccess -User $TemplateUser -ErrorAction Stop
}
catch{
    Write-Host "`t[x]:$TemplateUser couldn't be found." -ForegroundColor Red
    return
}

###########################################################


# Script : ################################################

if($Select){
    $FinalTemplate = $TemplateTeams | Out-GridView -Title "Select Channels" -PassThru
}else{
    $FinalTemplate = $TemplateTeams
}

Write-Host "`n`t(i):Attempting to copy rights." -ForegroundColor Cyan

# Adding user to the various Teams
$Teams = $FinalTemplate.GroupId | Select-Object -Unique
foreach($Team in $Teams){
    $param = @{
        GroupId = $Team
        User = $Identity
        ErrorAction = "Stop"
    }
    
    if($IncludeRights){
        $Channel = $FinalTemplate | Where-Object{$_.GroupId -eq $Team -and $_.Channel -eq "General"}
        if($Channel.Access -eq "Owner"){
            $param += @{Role = "Owner"}
        }
    }
    try{
        Add-TeamUser @param
    }
    catch{
        $TeamDisplayName = ($FinalTemplate | Where-Object{$_.GroupId -eq $Team})[0].Teams
        Write-Host "`t[x]:Failed to add user to Teams : $TeamDisplayName" -ForegroundColor Red
    }
}

Write-Host "`n`t(i):Adding private channel access." -ForegroundColor Cyan

# Adding user to the various channels
$Channels = $FinalTemplate | Where-Object{$_.Type -eq "Private"}
foreach($Channel in $Channels){
    $param = @{
        GroupId = $Channel.GroupId
        DisplayName = $Channel.Channel
        User = $Identity
        ErrorAction = "Stop"
    }
    
    if($IncludeRights){
        $Channel = $TemplateTeams | Where-Object{$_.GroupId -eq $Team -and $_.Channel -eq "General"}
        if($Channel.Access -eq "Owner"){
            $param += @{Role = "Owner"}
        }
    }
    try{
        Add-TeamChannelUser @param
    }
    catch{
        Write-Host "`t[x]:Failed to add user to channel : $($Channel.Channel)" -ForegroundColor Red
    }
}