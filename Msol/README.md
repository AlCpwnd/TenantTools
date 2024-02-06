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
