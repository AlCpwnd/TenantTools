param(
    [Parameter(Mandatory,Position=0)]
    [string]$sourceGroupId,
    [Parameter(Mandatory,Position=1)]
    [string]$destinationGroupId
)

$scopes = 'GroupMember.ReadWrite.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "(i) Scopes missing: $($scopes -join ',')" -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

$LogFile = $PSCommandPath.Replace('.ps1','.log')
if(-not (Test-Path -Path $LogFile)){
    New-Item -Path $LogFile
}

$logParams = @{
    FilePath = $LogFile
    Encoding = 'utf8'
    Append = $true
}

"`n$(Get-Date -Format "yyyy-MM-dd HH:mm") - Copy start" | Out-File @logParams
"Copying `"$((Get-MgGroup -GroupId $sourceGroupId).DisplayName)`" members to `"$((Get-MgGroup -GroupId $destinationGroupId).DisplayName)`"" | Out-File @logParams
"Source: $sourceGroupId","Destination: $destinationGroupId" | Out-File @logParams

$destinationMembers = Get-MgGroupMember -GroupId $destinationGroupId
$sourceMembers = Get-MgGroupMember -GroupId $sourceGroupId 
$missingMembers = Compare-Object $destinationMembers $sourceMembers -Property Id | Where-Object{$_.SideIndicator -eq '=>'}

if(-not $missingMembers){
    "! Given members are all already a part of the given Teams." | Out-File @logParams
    return
}

$i = 0
$iMax = $missingMembers.Count

$output = foreach($User in $missingMembers.Id){
    $i++
    Write-Progress -Activity "Copying permissions" -Status $User -PercentComplete (($i/$iMax)*100)
    try{
        New-MgGroupMember -GroupId $destinationGroupId -DirectoryObjectId $User -ErrorAction Stop
        "+ " + $groupMembers[$groupMembers.Id.IndexOf($User)].AdditionalProperties['displayName','mail'] -join "_"
    }catch{
        "! " + $groupMembers[$groupMembers.Id.IndexOf($User)].AdditionalProperties['displayName','mail'] -join "_"
    }
}
$output | Out-File @logParams
'=' * 30 | Out-File @logParams


<#
    .SYNOPSIS
    Adds members of the source group to the destination group.

    .DESCRIPTION
    Compares the members of the source group and will add the missing ones
    to the destination group. Throws an error if all the members are already
    member of the destination group.

    .PARAMETER sourceGroupId
    Object Id of the group you want the members to be added to the 
    destination group.

    .PARAMETER destinationGroupId
    Object Id of the group you want to copy the members to.

    .INPUTS
    None. You cannot pipe objects to GroupMemberCopy.ps1.

    .OUTPUTS
    The script will generate a log file containing the user copied over as
    well as the users which couldn't be copied over.

    .LINK
    Get-MgGroup

    .LINK
    Get-MgGroupMember

    .LINK
    New-MgGroupMember
#>