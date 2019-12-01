---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version: https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Get-CMHealthCheck.md
schema: 2.0.0
---

# Get-CMHealthCheck

## SYNOPSIS
Extract ConfigMgr Site data

## SYNTAX

```
Get-CMHealthCheck [[-SmsProvider] <String>] [[-OutputFolder] <String>] [[-NumberOfDays] <Int32>]
 [[-Healthcheckfilename] <String>] [-OverWrite] [-NoHotfix] [<CommonParameters>]
```

## DESCRIPTION
Exracts SCCM hierarchy and site server data
and stores the information in multiple XML data files which are then
processed using the Export-CM-Healthcheck.ps1 script to render
a final MS Word report.

## EXAMPLES

### EXAMPLE 1
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NumberofDays:30
```

### EXAMPLE 2
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -Overwrite -Verbose
```

### EXAMPLE 3
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -HealthcheckDebug -Verbose
```

### EXAMPLE 4
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NoHotfix
```

### EXAMPLE 5
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -OutputFolder "c:\temp"
```

### EXAMPLE 6
```
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -OutputFolder "c:\temp" -HealthcheckFilename ".\healthcheck.xml"
```

## PARAMETERS

### -SmsProvider
FQDN of SCCM site server

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: $($env:COMPUTERNAME)
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFolder
Path for output data files

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: "$($env:USERPROFILE)\Documents"
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberOfDays
Number of days to go back for alerts in logs (default = 7)

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: 7
Accept pipeline input: False
Accept wildcard characters: False
```

### -Healthcheckfilename
Name of configuration file (default is .\assets\cmhealthcheck.xml)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OverWrite
Overwrite existing output folder if found.
Folder is named by datestamp, so this only applies when
running repeatedly on the same date

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHotfix
Suppress hotfix inventory.
Can save significant runtime

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
* Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
* Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
* Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
* Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference
* Support: Database name must be CM_\<SITECODE\> (you need to adapt the queries if not this format)

* Security Rights: user running this tool should have the following rights:
- SQL Server (serveradmin) to be able to see database / cpu stats
- SCCM Database (db_owner) used to create/drop user-defined functions
- msdb Database (db_datareader) used to read backup information
- read-only analyst on the SCCM console
- local administrator on all computer (used to remotely connect to the registry and services)
- firewall allowing 1433 (or equivalent) to all SQL Servers (including SQL Express on Secondary Site)
- Remote WMI/DCOM firewall - http://msdn.microsoft.com/en-us/library/jj980508(v=winembedded.81).aspx
- Remote WUA - http://msdn.microsoft.com/en-us/library/windows/desktop/aa387288%28v=VS.85%29.aspx
- Comments: To get the free disk space, enable the Free Space (MB) for the Logical Disk

## RELATED LINKS

[https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Get-CMHealthCheck.md](https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Get-CMHealthCheck.md)

