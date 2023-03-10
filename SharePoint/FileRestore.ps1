param(
    [string]$User,
    [string]$Date
)

#Requires -Modules PnP.PowerShell

$RecyclyBinContents = Get-PnPRecycleBinItem
if($User){
    $UserFilteredContents = $RecyclyBinContents | Where-Object{$_.DeletedByEmail -eq $User}
    $RecyclyBinContents = $UserFilteredContents
    if(!$UserFilteredContents){
        throw "No deleted items found for given user: $User"
    }
}

if($Date){
    try{
        $Date = Get-Date $Date -ErrorAction Stop
        $DateFilteredContents = $RecyclyBinContents | Where-Object{$_.DeletedDate -le $Date}
        $RecyclyBinContents = $DateFilteredContents
    }catch{
        throw "Failed to conver $Date to DateTime variable"
    }
}

$i = 0
$iMax = $RecyclyBinContents.Count

$Global:Errors = foreach($Item in $RecyclyBinContents){
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
    Write-Host "Entrie stored in `'`$Global:Errors`' variable" @HighLIght
}