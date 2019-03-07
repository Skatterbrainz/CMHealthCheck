---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version:
schema: 2.0.0
---

# Export-CMHealthCheckHTML

## SYNOPSIS
Publish CMHealthCheck HTML Report

## SYNTAX

```
Export-CMHealthCheckHTML [-ReportFolder] <String> [[-OutputFolder] <String>] [-Detailed]
 [[-CustomerName] <String>] [-AutoConfig] [[-AuthorName] <String>] [[-CopyrightName] <String>]
 [[-Healthcheckfilename] <String>] [[-MessagesFilename] <String>] [[-Healthcheckdebug] <Object>]
 [[-Theme] <String>] [[-CssFilename] <String>] [[-TableRowStyle] <String>] [[-ImageFile] <String>] [-Overwrite]
 [<CommonParameters>]
```

## DESCRIPTION
Converts the data output from Get-CMHealthCheck to generate an HTML report file

## EXAMPLES

### EXAMPLE 1
```
Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein"
```

### EXAMPLE 2
```
Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
```

### EXAMPLE 3
```
Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -OutputFolder "c:\reports" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Theme 'Ocean' -TableRowStyle Dynamic -Verbose
```

### EXAMPLE 4
```
Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -AutoConfig -CustomerName "Contoso"
```

Applies custom parameters using "cmhealthconfig.txt" file in $env:USERPROFILE\Documents folder

## PARAMETERS

### -ReportFolder
Path to output data folder (e.g.
".\2019-03-06\cm01.contoso.local")

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFolder
Path to write new report file (default = User Documents folder)

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

### -CustomerName
Name of customer (default = "Customer Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
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

### -AuthorName
Report Author name (default = "Your Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: Your Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -CopyrightName
Text to use for copyright footer string (default = "Your Company Name"), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: Skatterbrainz
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
Position: 6
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
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Healthcheckdebug
Enable verbose output (or use -Verbose)

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
CSS style theme name, or 'Custom' to specify a file (default = 'Ocean'), or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: Ocean
Accept pipeline input: False
Accept wildcard characters: False
```

### -CssFilename
CSS file path to import when Theme is set to 'Custom', or use AutoConfig file

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

### -TableRowStyle
Apply CSS table style: Solid, Alternating, or Dynamic.
Default is Solid, or use AutoConfig file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 11
Default value: Solid
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageFile
Image Log file

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 12
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Overwrite
Overwrite existing report file if found, or use AutoConfig file

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
