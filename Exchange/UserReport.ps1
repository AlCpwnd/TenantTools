param(
    [ValidateScript({
        if($_ -notmatch ".csv"){
            throw "Output filetype must be 'csv'."
        }else{
            return $true
        }
    })]
    [System.IO.FileInfo]$OutputFile,
    [PSCredential]$Credentials
)

#Requires -Modules AzureAD,ExchangeOnlineManagement

if(!$OutputFile){
    $OutputFile = $PSCommandPath.Replace(".ps1",".csv")
}

if($Credentials){
    try{
        Connect-AzureAD -Credential $Credentials -ErrorAction Stop
        Connect-ExchangeOnline -Credential $Credentials -ErrorAction Stop
    }catch{
        throw "Failed to connect with the given credentials"
    }
}

$AzureUsers = Get-AzureADUser -Filter "UserType eq 'Member'" | Select-Object DisplayName,UserPrincipalName,DirSyncEnabled
$ExchangeMailboxes = Get-EXOMailbox | Select-Object DisplayName,UserPrincipalName,RecipientTypeDetails

class Office365Account{
    [String]$DisplayName
    [String]$Login
    [String]$MailboxType
    [ValidateSet('Sync','Cloud')][String]$Source
    Office365Account(
        [String]$d,
        [String]$l,
        [String]$m,
        [String]$s
    ){
        $this.DisplayName = $d
        $this.Login = $l
        $this.MailboxType = $m
        $this.Source = $s
    }
}

$Report = foreach($User in $AzureUsers){
    $MailAccount = $ExchangeMailboxes | Where-Object{$_.UserPrincipalName -eq $User.UserPrincipalName}
    if($User.DirSyncEnabled){
        $Source = 'Sync'
    }else{
        $Source = 'Cloud'
    }
    [Office365Account]::new($User.DisplayName,$User.UserPrincipalName,$MailAccount.RecipientTypeDetails,$Source)
}

$Report | Export-Csv -Path $OutputFile -NoTypeInformation -Delimiter ',' -Encoding utf8