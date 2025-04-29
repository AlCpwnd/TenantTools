# MSOL

These script are meanst for Microsoft license reporting.

---

## UserReport.ps1

### Synopsis

Returns existing users and their associated licenses.

### Syntax

```
UserReport.ps1 [-Export] [[-Path] <String>] [<CommonParameters>]
```

### Description

Return a CSV file containing all tenant members and the name of their associated licenses.

### Parameters

#### -Export

If enabled, will export the output into a CSV file.
```
Type: SwitchParameter
Parameter Sets: (All)

Required: false
Position: named
Default value: False
Accept pipeline: false
Accept wildcard characters: false
```

#### -Path

Path of the CSV file you want the data to be exported to.
```
Type: String
Parameter Sets: (All)

Required: false
Position: 1
Default value: None
Accept pipeline: false
Accept wildcard characters: false
```

### Related Links

* [Get-EXOMailbox](https://learn.microsoft.com/powershell/module/exchange/get-exomailbox)
* Get-MsolUser

---

## MgServiceCheck.ps1

### Synopsis

Returns a matrix of the users and licenses.

### Syntax

```
MgServiceCheck.ps1 [<CommonParameters>]
```

### Description

The script will attempt to connect to the Graph API using the
required scopes. If the current scopes aren't sufficient, it will
attempt to add the missing scropes to the current session.
Returns a matrix detailing which license are assined to which users.

### Related Links

* [Get-MgUser](https://learn.microsoft.com/powershell/module/microsoft.graph.users/get-mguser https://learn.microsoft.com/graph/api/intune-onboarding-user-get?view=graph-rest-1.0 https://learn.microsoft.com/graph/api/intune-mam-user-list?view=graph-rest-1.0)
* [Get-MgSubscribedSku](https://learn.microsoft.com/powershell/module/microsoft.graph.identity.directorymanagement/get-mgsubscribedsku https://learn.microsoft.com/graph/api/subscribedsku-get?view=graph-rest-1.0 https://learn.microsoft.com/graph/api/subscribedsku-list?view=graph-rest-1.0)
