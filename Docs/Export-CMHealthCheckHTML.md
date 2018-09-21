---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version:
schema: 2.0.0
---

# Export-CMHealthCheckHTML

## SYNOPSIS
Convert extracted ConfigMgr site information to HTML report

## SYNTAX

```
Export-CMHealthCheckHTML [-ReportFolder] <String> [[-OutputFolder] <String>] [-Detailed]
 [[-CustomerName] <String>] [[-AuthorName] <String>] [[-CopyrightName] <String>]
 [[-Healthcheckfilename] <String>] [[-MessagesFilename] <String>] [[-Healthcheckdebug] <Object>]
 [[-Theme] <String>] [[-CssFilename] <String>] [-DynamicTableRows] [-Overwrite] [<CommonParameters>]
```

## DESCRIPTION
Converts the data output from Get-CMHealthCheck to generate an HTML report file

## EXAMPLES

### EXAMPLE 1
```
Export-CMHealthCheckHTML -ReportFolder "2018-9-19\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein"
```

### EXAMPLE 2
```
Export-CMHealthCheckHTML -ReportFolder "2018-9-19\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
```

### EXAMPLE 3
```
Export-CMHealthCheckHTML -ReportFolder "2018-9-19\cm01.contoso.com" -OutputFolder "c:\reports" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Theme 'Ocean' -DynamicTableRows -Verbose
```

## PARAMETERS

### -ReportFolder
Path to output data folder (e.g.
".\2018-9-19\cm01.contoso.com")

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
Collect more granular data for final reporting

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
Name of customer (default = "Customer Name")

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

### -AuthorName
Report Author name (default = "Your Name")

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
Text to use for copyright footer string (default = "Your Company Name")

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
CSS style theme name, or 'Custom' to specify a file (default = 'Ocean')

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
CSS file path to import when Theme is set to 'Custom'

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

### -DynamicTableRows
Apply mouse-over dynamic table row effects

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

### -Overwrite
Overwrite existing report file if found

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
* 1.0.4 - 12/04/2017 - David Stein
* 1.0.5 - 09/19/2018 - David Stein
* Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
* Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
* Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
* Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference

## RELATED LINKS
