# Samples: AutoConfig File


## File: cmhealthconfig.txt

```
AuthorName=John Wick
CopyrightName=Retirement Specialists
Theme=Ocean
Detailed=True
TableRowStyle=Solid
CssFilename=c:\docs\disavowed.css
ImageFile=c:\docs\bodybags.png
CoverPage=
Template=
HealthcheckFilename=
MessagesFilename=
HealthcheckDebug=False
Overwrite=True
```

## Example Usage

```powershell
Export-CMHealthCheck -ReportFolder "2024-11-16\cm01.contoso.local" -AutoConfig -CustomerName "Contoso"
```
