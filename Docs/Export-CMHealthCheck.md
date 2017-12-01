---
external help file: CMHealthCheck-help.xml
Module Name: CMHealthCheck
online version: 
schema: 2.0.0
---

# Export-CMHealthCheck

## SYNOPSIS
Convert extracted ConfigMgr site information to Word Document

## SYNTAX

```
Export-CMHealthCheck [-ReportFolder] <String> [[-OutputFolder] <String>] [-Detailed] [[-CoverPage] <String>]
 [[-CustomerName] <String>] [[-AuthorName] <String>] [[-CopyrightName] <String>]
 [[-Healthcheckfilename] <String>] [[-MessagesFilename] <String>] [[-Healthcheckdebug] <Object>] [-Overwrite]
```

## DESCRIPTION
Converts the data output from Get-CMHealthCheck to generate a
report document using Microsoft Word (2010, 2013, 2016).
Intended
to be invoked on a desktop computer which has Office installed.

## EXAMPLES

### -------------------------- EXAMPLE 1 --------------------------
```
Export-CMHealthCheck -ReportFolder "2017-11-17\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
```

## PARAMETERS

### -ReportFolder
Path to output data folder (e.g.
".\2017-11-17\cm01.contoso.com")

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
Log folder path

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

### -CoverPage
Word theme cover page (default = "Slice (Light)")

```yaml
Type: String
Parameter Sets: (All)
Aliases: 

Required: False
Position: 3
Default value: Slice (Light)
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
Position: 4
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
Position: 5
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
Position: 6
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
Position: 7
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
Position: 8
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
Position: 9
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

## INPUTS

## OUTPUTS

## NOTES
* 1.0.3 - 12/01/2017 - David Stein
* Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
* Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
* Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
* Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference

## RELATED LINKS

