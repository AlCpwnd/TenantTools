$scopes = 'User.ReadWrite.All'

$MissingScopes = $scopes | Where-Object{(Get-MgContext).Scopes -notcontains $_}

if($MissingScopes){
    Write-Host "Scopes missing:$($scopes -join ',')"
    Connect-MgGraph -Scopes $MissingScopes
}

$Guests = Get-MgUser -All:$true -Filter "usertype eq 'Guest'" -Property Id,SurName,LastName,DisplayName,Mail | Where-Object{$_.DisplayName -match '\.|,'}

foreach($User in $Guests){
    if($User.DisplayName -match '\w\.\w'){
        $TextInfo = (Get-Culture).TextInfo
        $Temp = @{
            UserId = $User.Id
            DisplayName = $TextInfo.ToTitleCase($User.DisplayName.Replace('.',' '))
            GivenName = $TextInfo.ToTitleCase($User.DisplayName.Split('.')[0])
            SurName = $TextInfo.ToTitleCase($User.DisplayName.Split('.')[1])
        }
        if($Temp.Surname.Length -eq 1){
            Write-Host "ERROR: $($User.DisplayName)" -ForegroundColor Red
            continue
        }
    }elseif($User.DisplayName -match '\w\,\s'){
        $Name = $User.DisplayName.Split(',').Trim()
        $Temp = @{
            UserId = $User.Id
            DisplayName = $Name[1] + " " + $Name[0]
            GivenName = $Name[1]
            SurName = $Name[0]
        }
    }
    if($Temp.GivenName -match 'Dr\.'){
        $Temp.GivenName = $Temp.GivenName.Replace('Dr. ','')
    }
    if($Temp.DisplayName -match '\s\(.+\)\s'){
        $Bracket = $Temp.DisplayName.Split() | Where-Object{$_ -like "(*)"}
        $Temp.DisplayName = $Temp.DisplayName.Replace(" $Bracket",'') + " $Bracket"
        if($Temp.Surname -match $Bracket){
            $Temp.Surname = $Temp.Surname.Replace($Bracket,'').Trim()
        }elseif($Temp.GivenName -match $Bracket){
            $Temp.GivenName = $Temp.GivenName.Replace($Bracket,'').Trim()
        }
    }
    $Temp | Select-Object DisplayName,GivenName,SurName
}
