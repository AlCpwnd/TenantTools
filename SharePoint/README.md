# Sharepoint

Contains script meant to manage SharePoint.

---

## FileRestore.ps1

### Synopsis

Will restore files corresponding to the given filter variables.

### Syntax

```ps
FileRestore.ps1 [[-User] <string>] [[-Date] <string>]
```

### Description

The script will restore the items within the recyclebin that correspond to the given filtering variables.

### Examples

#### Example 1

```ps
FireRestore.ps1 -User john.doe@contosco.com
```

Restores all files that have been deleted by the user with the userprincipalname **<jogn.doe@contosco.com>**.

#### Example 2

```ps
FireRestore.ps1 -Date '31/12/2022'
```

Restores all files that have been deleted after **31/12/2022**.

#### Example 3

```ps
FireRestore.ps1 -User john.doe@contosco.com -Date '28/02/2023 09:00'
```

Restores all files that have been deleted after **31/12/2022 09:00**.

#### Example 4

```ps
FireRestore.ps1 -User john.doe@contosco.com -Date '01/04/2021 12:37'
```

Restores all files that have been deleted after **01/04/2021 12:37** by the user with the userprincipalname **<jogn.doe@contosco.com>**.

### Parameters

#### -User

Userprincipalname of email-address of the user who deleted the files that are to be restored.

```ps
Type: String
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

#### -Date

Date or DateTime after which the files have been deleted.

```ps
Type: String
Parameter Sets: (All)

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```
