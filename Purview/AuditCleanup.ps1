<#
    .SYNOPSIS
    Cleans up Microsoft Purview activity reports.

    .DESCRIPTION
    Takes in the activity report csv and cleans up the contents and columns in order for it to be
    more readable.

    .INPUTS
    The CSV file you download after a Microsoft Purview report has ran.

    .OUTPUTS
    If you have the 'ImportExcel' module install, the script will automatically generate a new
    Excel file containing the cleaned up date.
    If the module cannot be found, a new CSV file is generated next to the original with 
    "_cleaned" added to the file name.

    .PARAMETER LogFile
    CSV file report you downloaded from the Microsoft Purview portal.

    .LINK
    Import-Csv

    .LINK
    Export-Csv

    .LINK
    https://learn.microsoft.com/en-us/purview/audit-solutions-overview

    .EXAMPLE
    PS> .\AuditCleanup.ps1 -LogFile .\PurviewLogs.csv
    Cleaned up report exported to: .\PurviewLogs_cleaned.csv

#>


param(
    [Parameter(Mandatory)][string]$LogFile
)

try {
    $csv = Import-Csv -Path $LogFile -ErrorAction Stop
}
catch {
    Write-Host "Failed to import file: $LogFile" -ForegroundColor Red
    $Error[0]
    return
}

class FileReport {
    [string]$recordID
    [datetime]$creationDate
    [string]$operation
    [bool]$intuneManagedDevice
    [string]$deviceName
    [string]$userId
    [ipaddress]$clientIp
    [string]$browserName
    [string]$eventSource
    [string]$geoLocation
    [string]$itemType
    [string]$siteUrl
    [string]$relativeUrl
    [string]$file
    [string]$objectId
    FileReport(
        $Object
    ) {
        $auditData = $Object.AuditData | ConvertFrom-Json
        $this.recordID = $Object.RecordId
        $this.creationDate = $Object.CreationDate
        $this.operation = $Object.Operation
        $this.intuneManagedDevice = $auditData.IsManagedDevice
        $this.deviceName = $auditData.DeviceDisplayName
        $this.userId = $Object.UserId
        $this.clientIp = $auditData.clientIp
        $this.browserName = $auditData.BrowserName
        $this.eventSource = $auditData.EventSource
        $this.geoLocation = $auditData.GeoLocation
        $this.itemType = $auditData.ItemType
        $this.siteUrl = $auditData.SiteUrl
        $this.relativeUrl = $auditData.SourceRelativeUrl
        $this.file = $auditData.SourceFileName
        $this.objectId = $auditData.ObjectId
    }
}

$report = foreach ($record in $csv) {
    [FileReport]::new($record)
}

if (Get-Module -Name ImportExcel -ListAvailable -ErrorAction SilentlyContinue) {
    $report | Export-Excel -Show
}
else {
    $FilePath = $LogFile -replace '\.csv', '_cleaned.csv'
    $report | Export-Csv -Path $FilePath -Encoding utf8
    Write-Host "Cleaned up report exported to: $FilePath`n" -ForegroundColor Green
}