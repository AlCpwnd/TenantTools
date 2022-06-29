#Requires -Modules ExchangeOnlineManagement

<#
    .SYNOPSIS
    Configures a mail forward based on a csv input.

    .DESCRIPTION
    Reads the inputted csv file and configures the forward on the listed mailboxes to the listed email addresses.

    .PARAMETER File
    CSV file which contains the forward information. The file needs to have the following headers:
    DisplayName : displayname of the mailbox you want to configure the forward on.
    UserPrincipalName : UPN of the concerned mailbox.
    Mail : Email address toward which the mails need to be forwarded.

    .EXAMPLE
    PS> Forward.ps1 -File forward.csv

    .LINK
    Set-Mailbox

    .LINK
    Get-Mailbox

    .LINK
    Import-Csv
#>


param(
    [Parameter(Mandatory=$true,Position=0)]
    [String]$File
)

Write-Host "`n`t[i]Initiating script"


# Input check : #################################

try{ # Verifies the given path
    Test-Path -Path $File -ErrorAction Stop
    $csv = Import-Csv -Path $File
}
catch{
    Write-Host "`t[!]Invalid path : $File" -ForegroundColor Red
    return
}

# Verifies the CSV collumn headers
$Headers = $csv[0] | Get-Member -MemberType NoteProperty
$RequiredHeaders = "DisplayName","Mail","UserPrincipalName"
$MissingHeaders = @()

foreach($Header in $RequiredHeaders){ # Compares CSV headers with ones used in script
    if($Headers -notcontains $Header){
        $MissingHeaders += $Header
    }
}

if($MissingHeaders.Count -gt 0){ # If any are missing, aborts the script
    Write-Host "`t[!]Following data is missing from the file :" -ForegroundColor Red
    $MissingHeaders
    return
}

#################################################


# Script : ######################################

$i = 0
$iMax = $csv.Count

foreach($User in $csv){
    $i++
    Write-Progress -Activity "Configuring forward" -Status $User.DisplayName -PercentComplete (($i/$iMax)*100)
    
    try{
        Get-Mailbox $User.UserPrincipalName -ErrorAction Stop
    }
    catch{
        Write-Host "`t[*]$($User.UserPrincipalName) couldn't be found on the tenant"
        continue
    }

    if($User.Mail -notlike "*@*.*"){
        Write-Host "`t[!]$($User.DisplayName) doesn't have a valid forwarding mail : " -NoNewline -ForegroundColor Red
        $User.Mail
        continue
    }

    Set-Mailbox -Identity $User.UserPrincipalName -DeliverToMailboxAndForward $true -ForwardingSmtpAddress $User.Mail | Out-Null
    Write-Host "`t[i]Forward configured : " -NoNewline
    Write-Host "$($User.UserPrincipalName)" -ForegroundColor Yellow
}

#################################################
