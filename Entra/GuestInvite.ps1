#Requires -Modules Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Users

param(
    [Array]$email
)

$MgUsers = (Get-MgUser -Filter "userType eq 'Guest'" -All).Mail

$ExistingGuests = $Email | Select-Object -Unique | Where-Object{$MgUsers -contains $_}

$RemainingUsers = $MgUsers | Where-Object{$ExistingGuests -notcontains $_}

$PartialMatch = foreach($mail in $email){
    $MailStart = $mail.split("@")[0]
    $RemainingUsers | Where-Object{$_ -match $MailStart}
}
