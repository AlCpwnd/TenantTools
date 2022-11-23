param(
    [Parameter(Mandatory,Position=0)][ValidateSet("Standard","Detailed")][String]$Type
)

#Requires -Modules MsOnline

try{
    $Users = Get-MsolUser -EnabledFilter EnabledOnly -ErrorAction Stop | Where-Object{$_.IsLicensed}
}catch{
    Connect-MsolService
    $Users = Get-MsolUser -EnabledFilter EnabledOnly | Where-Object{$_.IsLicensed}
}

if($Type -eq "Standard"){
    $Users | Where-Object{$_.Licenses.ServiceStatus.ProvisioningStatus -contains "Disabled"}
    return
}

$Errors = foreach($User in $Users){
    foreach($License in $User.Licenses){
        if($License.ServiceStatus.ProvisioningStatus -contains "Disabled"){
        [PSCustomObject]@{
                DisplayName = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                License = $License.AccountSkuId.Split(":")[1]
                Service = ($License.ServiceStatus | Where-Object{$_.ProvisioningStatus -eq "Disabled"}).ServicePlan.ServiceName -Join ","
            }
        }
    }
}

if($Errors){
    Write-Host "`nService errors found: $($Errors.DisplayName.Count)"
    $Errors
}