#Requires -modules Microsoft.Graph.Teams,Microsoft.Graph.Users

param(
    [Parameter(Mandatory=$true,ParameterSetName='Single')]
    # Username or ID of the user you want to run a report for
    [String]$User,

    [Parameter(Mandatory=$true,ParameterSetName='All')]
    # Exports all existing Teams permissions
    [Switch]$All,

    [Parameter()]
    # File path you want to report to be written to
    [String]$Path,

    [Parameter()]
    # Will output the report as an Excel file
    [Switch]$Xlsx
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
    try{
        $UserInfo = Get-MgUser -UserId $User -ErrorAction Stop
    }catch{
        Write-Host "Could not find corresponding user on tenant: $User" -ForegroundColor Red
        return
    }
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
        $Members = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/teams/$($Team.Id)/channels/$($Channel.Id)/members").value
        if($PSCmdlet.ParameterSetName -eq "All"){
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
        }elseif($Members.userId -contains $UserInfo.Id){
            $UserPermission = $Members[$Members.UserId.IndexOf($UserInfo.Id)].Roles
            if($UserPermission){
                $Role = $UserPermission[0]
            }else{
                $Role = 'Member'
            }
            [TeamsPermission]::new(
                $Team.Id,
                $Team.DisplayName,
                $Channel.Id,
                $Channel.DisplayName,
                $Channel.MembershipType,
                $UserInfo.Mail,
                $Role
            )
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
        if($Path -match '\(\d+\).csv'){
            $Path = $Path -replace '\(\d+\).csv'," ($i).csv"
        }else{
            $Path = $Path.Replace(".csv"," ($i).csv")
        }
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
if($Xlsx){
    if(Get-Module -Name ImportExcel -ListAvailable){
        $Path = $Path.Replace('.csv','.xlsx')
        $Report | Export-Excel -Path $Path -WorksheetName Data -TableName Data -TableStyle Medium5 -IncludePivotTable -Show -PivotTableName Report -PivotRows Teams,Channel -PivotColumns Access -PivotData @{Access='Count'} -NoTotalsInPivot
    }else{
        Write-Host "The 'ImportExcel' module is required for this feature. Please install it and run the script again.`nDefaulting to CSV export." -ForegroundColor Red
        $Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
    }
}else{
    $Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation
}
Write-Host "File saved under: $Path" -ForegroundColor Green


<#
    .SYNOPSIS
    Generates a report of the existing Teams permissions.

    .DESCRIPTION
    Generates a report CSV or Excel report of either all existing Teams permissions or a single 
    user's Teams permissions.
    The report return a file containing the following information:
    - Group ID
    - Group DisplayName
    - Channel ID
    - Channel DisplayName
    - User's email address
    - User's permissions

    .NOTES
    The excel report export requires the module 'ImportExcel' to be installed on the device. If the
    option is selected and the module isn't present, the script will default to CSV and print a
    warning.
    
    .LINK
    Get-MgUser

    .LINK
    Get-MgUserJoinedTeams

    .LINK
    Get-MgTeam
    
    .EXAMPLE
    PS> MgAccessReport.ps1 -User john.doe@contoso.com -Path .\Report.csv
    
    .EXAMPLE
    PS> MgAccessReport.ps1 -User john.doe@contoso.com -Xlsx

    .EXAMPLE
    PS> MgAccessReport.ps1 -All
#>