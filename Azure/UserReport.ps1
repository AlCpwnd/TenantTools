#Requires -Modules Microsoft.Graph.Users,Microsoft.Graph.Groups

param(
    [Parameter(Mandatory, ParameterSetName = 'Group')][string]$GroupId,
    [Parameter(Mandatory, ParameterSetName = 'All')][switch]$All
)

Write-Host "(i) Setting up environment."

$scopes = 'User.Read.All', 'Group.Read.All'

$missingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($missingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $missingScopes
}

Write-Host "(i) Recovering users."

$properties = 'DisplayName', 'UserPrincipalName', 'SignInActivity', 'Mail'

try {
    switch ($PSCmdlet.ParameterSetName) {
        'All' { $users = Get-MgUser -All:$true -Property $properties -ErrorAction Stop }
        'Group' {
            $groupMembers = Get-MgGroupMember -GroupId $GroupId -All:$true -ErrorAction Stop
            $members = Get-MgUser -All:$true -Property $properties -ErrorAction Stop
            $i = 0
            $iMax = $members.Count
            $users = foreach ($member in $members) {
                Write-Progress -Activity "Recovering user details [$i/$iMax]" -Status $member.displayName -PercentComplete (($i / $iMax) * 100)
                $i++
                if($groupMembers.Id -contains $member.Id){
                    $member
                }
            }
            Write-Progress -Activity 'Documenting User' -Completed
        }
    }
}
catch {
    Write-Host "Failed to recover users." -ForegroundColor Red
    return "Failed to recover users."
}

Write-Host "(i) Generating users report."

$report = $users | Select-Object DisplayName, UserPrincipalName, Mail, @{l = 'LastSignIn'; e = { $_.SignInActivity.LastSignInDateTime } }

Write-Host "(i) Verifying report path."

$date = Get-Date -Format 'yyyyMMdd'

if ($All) {
    $fileName = "$date`_All.csv"
}
else {
    $fileName = "$date`_$GroupId.csv"
}

$filePath = (Get-Location).Path + "\$fileName"

$i = 1

while(Test-Path -Path $filePath){
    if($filePath -match '\(\d+\).csv'){
        $i++
        $filePath = $filePath -replace '\(\d+\).csv',"($i).csv"
    }else{
        $filePath = $filePath -replace '.csv'," ($i).csv"
    }
}

$report | Export-Csv -Path $filePath -NoTypeInformation -Encoding utf8

Write-Host "(i) Report exported to: $filePath`n"


<#
    .SYNOPSIS
    Recovers the requested users and makes a report containing their las login date.

    .DESCRIPTION
    The script will first setup the environment with the required permissions on your tenant in 
    order to be able to run. If you're prompted for credentials, confirm the requested permissions
    and approve them. 
    It then recovers the information of the given group members or all users and generates a report 
    containing their username and las logon date.

    .PARAMETER GroupId
    The script will return a report for the members of the given group.

    .PARAMETER All
    Will run the script for all users on the tenant.

    .OUTPUTS
    The script generates status messages and will generate a CSV file containing all the
    requested data.

    .EXAMPLE
    PS:> .\UserAccessReport.ps1 -GroupId 6605147e-040c-426e-ae3f-30c8fd41b4c2

    .EXAMPLE
    PS:> .\UserAccessReport.ps1 -All

    .LINK
    Connect-MgGraph

    .LINK
    Get-MgContext
    
    .LINK
    Get-MgUser

    .LINK
    Get-MgGroupMember
#>