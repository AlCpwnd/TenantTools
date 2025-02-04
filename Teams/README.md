# Teams Scripts

Scripts meant for managing and reporting Teams within the tenant.

The scripts starting with "Mg" will use the Microsoft Graph API and might be faster depending on the required change or reporting.

---

## AccessCopy.ps1

### Synopsis

Copies a user's Teams permissions onto another.

### Syntax

```
C:\Users\ADJ\Git\TenantTools\Teams\AccessCopy.ps1 [-Template] <String> [-Identity] <String> [-IncludeRights] [-Select] [<CommonParameters>]
```

### Description

Replicates a user's Teams channel membership onto another user.

### Examples

#### Example 1

```ps
AccessCopy.ps1 -Template 'j.smith@contosco.com' -Identity 'j.doe@contosco.com'
```

Copies the permissions of user **<j.smith@contosco.com>** onto **<j.doe@contosco.com>**.

#### Example 2

```ps
AccessCopy.ps1 -Template 'j.smith@contosco.com' -Identity 'j.doe@contosco.com -IncludeRight
```

Copies the permissions of user **<j.smith@contosco.com>** onto **<j.doe@contosco.com>**, including Teams channel ownership.

#### Example 3

```ps
AccessCopy.ps1 -Template '.\Permissions.csv' -Identity 'j.doe@contosco.com -Select
```

Extracts the permissions detailed in `.\Permissions.csv`, presents them to them to the user ([Select](#Select)), before applying them to the user **<j.doe@contosco.com>**.

### Parameters

#### -Template

UserPrincipalName of the user you want to replicate the permissions of.
This can alse be the path to a CSV file containing the template user's permissions. The CSV would have to been generated using the AccessReport.ps1 script.

```ps
Type: String
Parameter Sets: (All)

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

#### -Identity

User on which the permissions are to be applied.

```ps
Type: String
Parameter Sets: (All)

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

#### -IncludeRights

Will replicate the exact permissions the user. If not mentionned, the given user will be added as member or guest to each Teams.

```ps
Type: Switch
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

#### -Select

Will prompt you with the existing permissions and will ask you to point out which ones you wish to copy over.

```ps
Type: Switch
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

---

## AccessReport.ps1

### Synopsis

Documents permissions of a user.

### Syntax

```
AccessReport.ps1 [-User] <String> [[-Path] <String>] [<CommonParameters>]
```

### Description

Returns a CSV containing the Teams, Channels and the accesslevel the user has on those.

### Examples

#### Example 1

```ps
AccessReport.ps1 -User 'john.doe@contosco.com'
```

Return the existing Teams channel access and permissions for the user **<jogn.doe@contosco.com>**.

#### Example 2

```ps
AccessReport.ps1 -User 'john.doe@contosco.com' -Path '.\Report.csv
```

Return the existing Teams channel access and permissions for the user **<jogn.doe@contosco.com>** and exports the report to `.\Report.csv`.

### Parameters

#### -User

Userprincipalname of email-address of the user for which you want to generate the report.

```ps
Type: String
Parameter Sets: (All)

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

#### -Path

Filepath on which the report will be generated.

```ps
Type: String
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

---

## MgTeamsStructure.ps1

### Synopsis

Reports the current Teams and Channels structure.

### Syntax

```
MgTeamsStructure.ps1 [<CommonParameters>]
```

### Description

Returns a CSV containing the Teams, their access type, the Channel 
and their access type. Using the Microsoft Graph API

### Examples

#### Example 1

```ps
MgTeamsStructure.ps1
```

### Related Links

* [Get-MgTeam](https://learn.microsoft.com/powershell/module/microsoft.graph.teams/get-mgteam
* [Get-MgTeamChannel](https://learn.microsoft.com/powershell/module/microsoft.graph.teams/get-mgteamchannel
