---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version:
schema: 2.0.0
---

# Export-CMHealthReport

## SYNOPSIS
Convert extracted ConfigMgr site information to Word Document

## SYNTAX

```
Export-CMHealthReport [[-ReportType] <String>] [-ReportFolder] <String> [[-OutputFolder] <String>]
 [[-CustomerName] <String>] [-AutoConfig] [[-SmsProvider] <String>] [-Detailed] [[-CoverPage] <String>]
 [[-Template] <String>] [[-AuthorName] <String>] [[-CopyrightName] <String>] [[-Healthcheckfilename] <String>]
 [[-MessagesFilename] <String>] [[-Healthcheckdebug] <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
Converts the data output from Get-CMHealthCheck to generate a
report document using Microsoft Word (2010, 2013, 2016).
Intended
to be invoked on a desktop computer which has Office installed.

## EXAMPLES

### EXAMPLE 1
```
Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
```

### EXAMPLE 2
```
Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -Template ".\contoso.docx" -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
```

### EXAMPLE 3
```
Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -AutoConfig -CustomerName "Contoso"
```

## PARAMETERS

### -ReportType
Report output type (HTML or MS Word)

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Word
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReportFolder
Path to output data folder (e.g.
".\2019-03-06\cm01.contoso.local")

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFolder
Log folder path

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: "$($env:USERPROFILE)\Documents"
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
Position: 4
Default value: Customer Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoConfig
Use an auto configuration file, cmhealthconfig.txt in $env:USERPROFILE\documents folder
to fill-in AuthorName, CopyrightName, Theme, CssFilename, TableRowStyle

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

### -SmsProvider
{{Fill SmsProvider Description}}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Detailed
Collect more granular data for final reporting, or use AutoConfig file

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

### -CoverPage
Word theme cover page (default = "Slice (Light)"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: Slice (Light)
Accept pipeline input: False
Accept wildcard characters: False
```

### -Template
Word document file to use as a template.
Should have a cover page already in place.
If Template is specified, CoverPage and Copyright are ignored.

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

### -AuthorName
Report Author name (default = "Your Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
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
Position: 9
Default value: Your Company Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -Healthcheckfilename
Healthcheck configuration XML file name (default = ".\assets\cmhealthcheck.xml")
The file can be local, UNC or URI sourced as well

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

### -Healthcheckdebug
Enable verbose output (or use -Verbose)

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable.
For more information, see about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
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

## RELATED LINKS
