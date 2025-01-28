#Requires -modules Microsoft.Graph.Identity.DirectoryManagement,Microsoft.Graph.Users

$scopes = 'LicenseAssignment.Read.All','User.Read.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "Scopes missing:$($scopes -join ',')"
    Connect-MgGraph -Scopes $MissingScopes
}

Write-Host "Recovering license types."
$Licenses = Get-MgSubscribedSku | Where-Object{$_.AppliesTo -eq 'User' -and $_.CapabilityStatus -ne 'Suspended' -and $_.ConsumedUnits}

class UserLicense {
    [String]$ObjectId
    [String]$DisplayName
    [String]$Mail
    [String]$Type
    UserLincense(
        [String]$o,
        [String]$d,
        [String]$m,
        [String]$t
    ){
        $this.ObjectId = $o
        $this.DisplayName = $d
        $this.Mail = $m
        $this.Type = $t
    }
}

Write-Host "Recovering users."
$Users = Get-MgUser -All:$true

$i = 1
$iMax = $Users.Count
$Report = foreach($User in $Users){
    Write-Progress -Activity "Documenting Users[$i/$iMax]" -Status $User.DisplayName -PercentComplete (($i/$iMax)*100)
    $i++
    $UserInfo = Get-MgUser -UserId $User.Id -Property assignedLicenses,userType
    if($UserInfo.AssignedLicenses.Count -eq 0){
        continue
    }
    $Temp = [UserLicense]::new()
    $Temp.ObjectId = $User.Id
    $Temp.DisplayName = $User.DisplayName
    $Temp.Mail = $User.Mail
    $Temp.Type = $UserInfo.UserType
    foreach($License in $Licenses){
        if($UserInfo.AssignedLicenses.SkuId -contains $License.SkuId){
            $Test = $true
        }else{
            $Test = $false
        }
        $Temp | Add-Member -NotePropertyName $License.SkuPartNumber -NotePropertyValue $Test
    }
    $Temp
}
Write-Progress -Activity "Documenting Users" -Completed

$Path = "$PSScriptRoot\$(Get-Date -Format yyyyMMdd)_LicenseReport.csv" 
$i = 1
while(Test-Path -Path $Path){
    if($Path -like "*(*).csv"){
        $i ++
        $Path = $Path.Replace("($($i-1)).csv","($i).csv")
    }else{
        $Path = $Path.Replace(".csv"," ($i).csv")
    }
}

Write-Host "File saved under: $Path"
$Report | Export-Csv -Path $Path -Encoding UTF8 -NoTypeInformation

<#
    .SYNOPSIS
    Returns a matrix of the users and licenses.

    .DESCRIPTION
    The script will attempt to connect to the Graph API using the
    required scopes. If the current scopes aren't sufficient, it will
    attempt to add the missing scropes to the current session.
    Returns a matrix detailing which license are assined to which users.

    .INPUTS
    None. You cannot pipe objects to MgServiceCheck.ps1.

    .OUTPUTS
    A CSV file named using the following template: 
    <Date formatted yyyyMMdd>_LicenseReport.csv
    This file will be outputted in the current script directory.

    .LINK
    Get-MgUser

    .LINK
    Get-MgSubscribedSku
#>