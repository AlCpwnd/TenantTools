#Requires -Modules Microsoft.Graph.Users

[CmdletBinding()]

param(
    [Parameter()]
    # Run the script in test mode. The result of the changes will be displayed on the host and no changes will be made to the existing users.
    [Switch]$Test
)

$scopes = 'User.ReadWrite.All'

$missingScopes = $scopes | Where-Object { (Get-MgContext).Scopes -notcontains $_ }

if ($missingScopes) {
    Write-Host "The following scopes are missing:" -ForegroundColor Yellow
    $missingScopes | ForEach-Object { Write-Host "`t> $_" -ForegroundColor Yellow }
    Connect-MgGraph -Scopes $MissingScopes
}

$guests = Get-MgUser -Filter "userType eq 'Guest'" -All -Property DisplayName, Id, Mail, GivenName, SurName | Select-Object @{l = 'UserId'; e = { $_.id } }, displayName, mail, givenName, surName

function Test-UserFields {
    [OutputType([Bool])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        # User to run the tests for.
        [PSCustomObject]$User
    )
    $output = $true
    
    $properties = 'displayName', 'givenName', 'surName'

    foreach ($property in $properties) {
        if ($User.$property -cmatch '[A-Z]{2,}' -or -not $User.$property) {
            $output = $false
        }
    }

    return $output

    <#
    .SYNOPSIS
    Tests the given user for missing properties.
    
    .DESCRIPTION
    Goes over the following properties of the given object and return if any of them are empty of written fully in capital letters:
    > DisplayName
    > Given Name
    > Surname

    Returns "$true" if the user passes the tests. Returns "$false" otherwise.

    .INPUTS
    The user object with the appropriate properties.

    .OUTPUTS
    Returns a bool depending on if the tests are successful if not.
    #>
}

$TextInfo = (Get-Culture).TextInfo

if ($test) {
    $testCase = @()
}

foreach ($guest in $guests) {
    if (Test-UserFields -User $guest) {
        continue
    }
    if (-not($guest.givenName -and $guest.surName)) {
        if ($guest.displayName -match ',') {
            $split = $guest.DisplayName.Split(',')
            $guest.givenName = $split[1].Trim()
            $guest.surName = $split[0].Replace(',', '').Trim()
        }
        else {
            $split = $guest.DisplayName.Split()
            $guest.givenName = $split[0].Trim()
            $guest.surName = ($split[1..($split.Length - 1)] -join ' ').Trim()
        }
    }
    $properties = 'displayName', 'givenName', 'surName'
    foreach ($property in $properties) {
        if ($guest.$property -cmatch '[A-Z]{2,}') {
            $guest.$property = $TextInfo.ToTitleCase($guest.$property.ToLower())
        }
        if ($guest.$property -match '  ') {
            $guest.$property = $guest.$property.Replace('  ', ' ')
        }
    }
    if ($test) {
        $testCase += $guest
    }
    else {
        $params = @{}
        $guest.PSObject.Properties | ForEach-Object { $params[$_.Name] = $_.Value }
        try {
            Update-MgUser @params -ErrorAction Stop
        }
        catch {
            Write-Host "Failed to update: $($guest.DisplayName)" -ForegroundColor Red
            $guest
        }
    }
}

if ($Test) {
    return $testCase
}


<#
.SYNOPSIS
Goes over guests and complete any missing surname or given name.

.DESCRIPTION
Recovers all existing guests on the tenant and checks for:
> Missing surname
> Missing givenName
> Surname, givenName or displayName fully written in capital letters.

Correct the or completes them accordingly. The test switch can be used to have a visualization of the planned changes.

.INPUTS
None.

.OUTPUTS
If the "Test" switch is selected, a list of the expected changes.

.LINK
Get-MgUser

.LINK
Update-MgUser

.EXAMPLE
PS> .\GuestCompletion.ps1 -Test
UserId      : 3f52ac44-3c61-4692-800d-6d3bab629775                                         
DisplayName : Doe, John
Mail        : john.doe@contoso.com
GivenName   : John
Surname     : Doe

.EXAMPLE
PS> .\GuestCompletion.ps1
#>

