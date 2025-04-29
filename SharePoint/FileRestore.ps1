param(
    [string]$User,
    [string]$Date
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