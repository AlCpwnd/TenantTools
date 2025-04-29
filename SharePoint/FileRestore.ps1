param(
    [Parameter(ParameterSetName='User')]
    [string]$User,
    [Parameter(ParameterSetName='Date')]
    [string]$Date,
    [Parameter(ParameterSetName='Date')]
    [Parameter(ParameterSetName='User')]
    [switch]$ReportOnly
)

#Requires -Modules PnP.PowerShell

$RecycleBinContents = Get-PnPRecycleBinItem
if($User){
    $UserFilteredContents = $RecycleBinContents | Where-Object{$_.DeletedByEmail -eq $User}
    $RecycleBinContents = $UserFilteredContents
    if(!$UserFilteredContents){
        throw "No deleted items found for given user: $User"
    }
}

if($Date){
    try{
        $Date = Get-Date $Date -ErrorAction Stop
        $DateFilteredContents = $RecycleBinContents | Where-Object{$_.DeletedDate -le $Date}
        $RecycleBinContents = $DateFilteredContents
    }catch{
        throw "Failed to convert $Date to DateTime variable"
    }
}

if($ReportOnly){
    Write-Host "Files available for restoring "
    $RecycleBinContents | Select-Object DeletedDate,DeletedByName,ItemType,LeafName
    return
}

$i = 0
$iMax = $RecycleBinContents.Count

$Global:Errors = foreach($Item in $RecycleBinContents){
    Write-Progress -Activity 'Restoring' -Status $Item.Title -PercentComplete (($i/$iMax)*100)
    try{
        Get-PnPRecycleBinItem -Identity $Item.Id | Restore-PnPRecycleBinItem -Force -ErrorAction Stop
    }catch{
        $Item | Select-Object Title,Dirname
    }
    $i++
}

if($Errors){
    $HighLIght = @{
        ForegroundColor = 'Black'
        BackgroundColor = 'Red'
    }
    Write-Host 'Failed to restore the following items:' @HighLIght
    $Global:Errors
    Write-Host "Entries stored in `'`$Global:Errors`' variable" @HighLIght
}


<#
    .SYNOPSIS
    Copies a user's Teams permissions onto another.

    .DESCRIPTION
    Replicates a user's Teams channel membership onto another user.

    .PARAMETER User
    UserPrincipalName of the user who deleted the files.

    .PARAMETER Date
    Date at which the files were removed.

    .PARAMETER ReportOnly
    Will only list the file

    .INPUTS
    None. You cannot pipe objects into FileRestore.ps1 .

    .OUTPUTS
    None.

    .EXAMPLE
    >PS FileRestore.ps1 -User 'j.smith@contosco.com'

    .EXAMPLE
    >PS FileRestore.ps1 -Date '27/07/2006'

    .LINK
    Get-PnPRecycleBinItem

    .LINK
    Restore-PnPRecycleBinItem
#>