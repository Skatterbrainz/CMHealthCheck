---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version: https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Invoke-CMHealthCheck.md
schema: 2.0.0
---

# Invoke-CMHealthCheck

## SYNOPSIS
Generate Health Information from a Configuration Manager site

## SYNTAX

```
Invoke-CMHealthCheck [[-SmsProvider] <String>] [[-CustomerName] <String>] [[-AuthorName] <String>]
 [[-CopyrightName] <String>] [[-DataFolder] <String>] [[-PublishFolder] <String>] [-OpenBrowser] [-OverWrite]
 [-NoHotfix] [-AutoConfig] [-Detailed] [[-Template] <String>] [[-ReportType] <String>]
 [[-NumberOfDays] <Int32>] [-Healthcheckdebug] [[-Healthcheckfilename] <String>] [[-MessagesFilename] <String>]
 [<CommonParameters>]
```

## DESCRIPTION
Generate Health Information from a Configuration Manager site

## EXAMPLES

### EXAMPLE 1
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp"
```

Standard/default settings to collect data, and generate HTML report

### EXAMPLE 2
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -Overwrite
```

Replaces an existing (previous) output from the same date

### EXAMPLE 3
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -OpenBrowser
```

Opens the HTML report in the default web browser, upon completion

### EXAMPLE 4
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -Detailed
```

Generates additional detail in the output report file

### EXAMPLE 5
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -AutoConfig "config.txt"
```

Loads reporting parameters from custom text file

### EXAMPLE 6
```
Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -NoHotFix
```

Skips inventory of installed operating system hotfixes

## PARAMETERS

### -SmsProvider
FQDN of the SMS Provider host in the Configuration Manager site

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: "$(($env:COMPUTERNAME, $env:USERDNSDOMAIN) -join '.')"
Accept pipeline input: False
Accept wildcard characters: False
```

### -CustomerName
Name of customer (default = "Customer Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Customer Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -AuthorName
Report Author name (default = "Your Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Your Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -CopyrightName
Text to use for copyright footer string (default = "Your Company Name")

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: Your Company Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataFolder
Path to output data for storing files during collection phase

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: "$($env:USERPROFILE)\Documents"
Accept pipeline input: False
Accept wildcard characters: False
```

### -PublishFolder
Path to save the HTML report file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: "$($env:USERPROFILE)\Documents"
Accept pipeline input: False
Accept wildcard characters: False
```

### -OpenBrowser
Open HTML report in default web browser upon completion

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
Skip inventory of installed hotfixes

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

### -AutoConfig
Load parameters from configuration file
Example:
	\`\`\`
	Sample AutoConfig file cmhealthconfig.txt...
	AuthorName=John Wick
	CopyrightName=Retirement Specialists
	Theme=Ocean
	Detailed=True
	TableRowStyle=Solid
	CssFilename=c:\docs\wickrocks.css
	ImageFile=c:\docs\bodybags.png
	CoverPage=
	Template=
	HealthcheckFilename=
	MessagesFilename=
	HealthcheckDebug=False
	Overwrite=True
	\`\`\`

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

### -Detailed
Display additional details (verbose)

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

### -Template
{{ Fill Template Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReportType
{{ Fill ReportType Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: HTML
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
Position: 9
Default value: 7
Accept pipeline input: False
Accept wildcard characters: False
```

### -Healthcheckdebug
Enable verbose output (or use -Verbose)

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

### -Healthcheckfilename
Name of configuration file (default is .\assets\cmhealthcheck.xml)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MessagesFilename
Status and error message lookup table (default = ".\assets\messages.xml")
The file can be local, UNC or URI sourced as well

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### collected data files, folders, and output HTML file, happiness, confusion, consternation, whatever that means
## NOTES
New function for 1.0.11 - 10/2019

## RELATED LINKS

[https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Invoke-CMHealthCheck.md](https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Invoke-CMHealthCheck.md)

