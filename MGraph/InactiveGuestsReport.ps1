#Requires -modules Microsoft.Graph.Users

param(
    [parameter(Mandatory = $true)]
    [ValidateRange(1, [int]::MaxValue)]
    # Number of months you after which you would consider a user inactive
    [int]$number = 6,

    [Parameter(Mandatory = $false)]
    [ValidateSet('xlsx', 'csv')]
    # File format you want the report to be outputted in
    [String]$FileFormat
)

function Test-FilePath {
    param(
        [Parameter(Mandatory=$true)]
        # File path that needs to be verified.
        [System.IO.FileInfo]$Path
    )
    $i = 1
    while(Test-Path -Path $Path){
        if($Path -match '\(\d+\)\.'){
            $Path = $Path -replace '\(\d+\)', "($i)"
        }else{
            $Path = $Path -replace '\.(?:[0-9a-zA-Z]+)$', "($i)$($Matches[0])"
        }
    }
    return $Path.FullName
}

$scopes = 'User.Read.All'

$MissingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($MissingScopes) {
    Write-Host "Scopes missing:" -ForegroundColor Yellow
    $scopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Write-Host "Adding scopes to current environment. Please allow the connection." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $MissingScopes
}

$Users = Get-MgUser -Filter "userType eq 'Guest'" -All:$true -Property DisplayName, mail, AccountEnabled, SignInActivity, ExternalUserState, CreatedDateTime, Sponsors -ExpandProperty Sponsors

$limit = (Get-Date).AddMonths(-$number)

$report = $Users | Where-Object { $_.SignInActivity.LastSignInDateTime -lt $limit -and $_.CreatedDateTime -lt $limit } | Select-Object Id, DisplayName, Mail, @{l = 'Blocked'; e = { $_.AccountEnabled } }, @{ l = 'LastSignIn'; e = { $_.SignInActivity.LastSignInDateTime } }, ExternalUserState, CreatedDateTime, @{ l = 'Sponsor'; e = { $_.Sponsors.AdditionalProperties.userPrincipalName } }

$date = Get-Date -Format yyyyMMdd

if ($FileFormat -eq 'xlsx') {
    if (Get-Module -Name ImportExcel -ListAvailable) {
        $fileName = (Get-Location).Path + "\{0}_InactiveUsers_{1}Months.xlsx" -f $date, $number
        $fileName = Test-FilePath -Path $fileName
        $report | Export-Excel -Path $fileName -ClearSheet -WorksheetName Report -TableName ActivityReport
    }
    else {
        Write-Host "The module 'ImportExcel' wasn't found on the device. This module is required for exporting in this format. Defaulting to CSV." -ForegroundColor Red
        $fileName = (Get-Location).Path + "\{0}_InactiveUsers_{1}Months.csv" -f $date, $number
        $fileName = Test-FilePath -Path $fileName
        $report | Export-Csv -Path $fileName -Encoding utf8
    }
}
else {
    $fileName = (Get-Location).Path + "\{0}_InactiveUsers_{1}Months.csv" -f $date, $number
    $fileName = Test-FilePath -Path $fileName
    $report | Export-Csv -Path $fileName -Encoding utf8
}

Write-Host "Report exported to: $fileName" -ForegroundColor Green

<#
    .SYNOPSIS
    Return guest users that haven't signed in within the given period.

    .DESCRIPTION
    Returns a list of guest users that haven't signed in on the tenant for given amount of months.(Default is 6 months)
    And exports it in the requested file format.

    .NOTES
    Excel file export requires the "ImportExcel" module to be installed.

    .LINK
    Get-MgUser
#>

