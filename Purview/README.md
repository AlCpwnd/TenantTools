# :mag: Purview

This library contains scripts to interact with the files and services of the Microsoft Purview solution.

---

## AuditCleanup.ps1

### Synopsis

Cleans up Microsoft Purview activity reports.

### Syntax

```TXT
AuditCleanup.ps1 [-LogFile] <String> [<CommonParameters>]
```

### Description

Takes in the activity report csv and cleans up the contents and columns in order for it to be
more readable.

### Examples

#### Example 1

```ps
.\AuditCleanup.ps1 -LogFile .\PurviewLogs.csv
```

### Parameters

#### -LogFile

CSV file report you downloaded from the Microsoft Purview portal.

```TXT
Type: String

Required: true
Position: 1
Default value: None
Accept pipeline: false
Accept wildcard characters: false
```

### Related Links

- [Import-Csv](https://go.microsoft.com/fwlink/?LinkID=113341)

- [Export-Csv](https://go.microsoft.com/fwlink/?LinkID=113299)

