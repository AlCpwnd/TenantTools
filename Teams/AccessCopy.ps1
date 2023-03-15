#Requires -Modules MicrosoftTeams

param(
    [Parameter(Mandatory = $true)]
    [String]$Template,
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

    class TeamsPermission{
        [String]$GroupId
        [String]$Teams
        [String]$Channel
        [String]$Access
        TeamsPermission(
            [String]$g,
            [String]$t,
            [String]$c,
            [String]$a
        ){
            $this.GroupId = $g
            $this.Teams = $t
            $this.Channel = $c
            $this.Access = $a
        }
    }

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
                [TeamsPermission]::new($Team.GroupId,$Team.DisplayName,$Channel.DisplayName,$UserInfo)
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

function Write-Info{Param([Parameter(Mandatory,Position=0)][string]$Message)Write-Host "`t(i):$Message" -ForegroundColor Cyan}

function Write-Warning{Param([Parameter(Mandatory,Position=0)][string]$Message)Write-Host "`t[!]:$Message" -ForegroundColor Red}

###########################################################


# Input Verification : ####################################

# Lists the $Identity user's existing Teams access.
try{
    Write-Info "Verifying [$Identity] current access."
    Get-Team -User $Identity -ErrorAction Stop
}
catch{
    throw "$Identity couldn't be found."
}

# Lists the $Template user's existing Teams access.
if($Template -like "*@*.*"){
    try{
        Write-Info "Recovering [$Template] current access."
        $TemplateTeams = Get-TeamsChannelAccess -User $Template -ErrorAction Stop
    }
    catch{
        throw "$Template couldn't be found."
    }
}elseif (Test-Path -Path $Template -PathType Leaf) {
    try{
        $TemplateTeams = Import-Csv -Path $Template -Encoding UTF8 -ErrorAction Stop
    }catch{
        throw "Failed to import: $Template"
    }
}else{
    throw 'Invalid Template parameter.'
}

###########################################################


# Script : ################################################

if($Select){
    $FinalTemplate = $TemplateTeams | Out-GridView -Title "Select Channels" -PassThru
}else{
    $FinalTemplate = $TemplateTeams
}

Write-Info "Attempting to copy rights."

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
        Write-Error "Failed to add user to Teams : $TeamDisplayName"
    }
}

Write-Info "Adding private channel access."

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
        Write-Error "Failed to add user to channel : $($Channel.Channel)"
    }
}

<#
    .SYNOPSIS

    Copies a user's Teams permissions onto another.

    .DESCRIPTION

    Replicates a user's Teams channel membership onto another user.

    .PARAMETER Template

    UserPrincipalName of the user you want to replicate the permissions of.
    This can alse be a CSV file containing the template user's permissions. The CSV would have to been generated using the AccessReport.ps1 script.

    .PARAMETER Identity

    User on which the permissions are to be applied.

    .PARAMETER IncludeRights

    Will replicate the exact permissions the user. If not mentionned, the given user will be added as member or guest to each Teams.

    .PARAMETER Select

    Will prompt you with the existing permissions and will ask you to point out which ones you wish to copy over.

    .INPUTS

    None. You cannot pipe objects into AccessCopy.ps1 .

    .OUTPUTS

    None.

    .LINK

    Get-Team

    .LINK

    Get-TeamChannelUser

    .LINK

    Get-TeamChannel

    .LINK

    Add-TeamUser

    .LINK

    Add-TeamChannelUser
#>