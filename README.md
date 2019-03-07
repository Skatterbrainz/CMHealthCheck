# CMHealthCheck

ConfigMgr Health-Check Reporting PowerShell functions

## Summary

CMHealthCheck is intended to make the former standalone PowerShell scripts ("get-cm-inventory.ps1" and "export-cm-healthcheck.ps1") easier to invoke as a module from PowerShell Gallery.  Both scripts have been refactored into a simple PowerShell module, with two public functions: Get-CMHealthCheck and Export-CMHealthCheck.  The required support XML data files are now invoked via URI from GitHub gist locations (on this account), but they can be downloaded for offline use as well.

PowerShell modules require PowerShell version 3.0 or later.  This was tested mostly on PowerShell 5.0 and 5.1.

## Installation and Usage

_Part 1 - Data Collection_

1. Log onto SCCM site server with full admin credentials (SCCM, Windows and SQL Server instance)
2. Open PowerShell console using Run as Administrator
3. Enter: Install-Module CMHealthCheck
4. Run: Get-CMHealthCheck (parameters...)
5. Collect output files from $env:USERPROFILE\Documents\YYYY-MM-DD\hostname
6. Copy to machine which has Office 2013 or 2016 installed (part 2)

_Part 2 - Reporting_

1. Log onto a Windows computer which has Office 2013 or 2016 installed
2. Open PowerShell console using Run as Administrator
3. Enter: Install-Module CMHealthCheck
4. Run: Export-CMHeathCheck (parameters...)
5. Wait for Document to finish building, Save document
6. Review report, add comments, dance around, drink, run outside buck nekkid and laugh hysterically

_Part 2a - Reporting_

1. Same as Part 2, but...
2. Run: Export-CMHealthCheckHTML (parameters...)
3. Open the HTML output file in your browser
4. Stare in awe as you mindlessly send me all of your earnings, ok, maybe not
5. Edit the report, add comments, laugh, snort, cough, sleep

## Notes

* Refer to the markdown files under \Docs for syntax and examples.
* Tested with the following environments:
   * ConfigMgr 2012 R2, 1610 to 1902.2 (GA and Tech Preview)
   * Windows Server 2012, 2012 R2, 2016
   * SQL Server 2012, 2014, 2016, 2017
   * Windows 10 1703 to 1809
   * Office 2013, 2016 / 365
   * PowerShell 3.0, 5.0, 5.1
  
* If you like it, please share with others.  If you hate it, tell me why so I can improve on it?
* Please submit bugs, comments, requests via the "Issues" link above.

## Removal and Cleanup

* To remove CMHealthCheck module and related files, use the Remove-Module or Uninstall-Module cmdlets.

## Examples

### Example 1

Installation and execution on ConfigMgr primary site server cm01.contoso.local site P01...

```powershell
Install-Module CMHealthCheck
.\Get-CMHealthCheck -SmsProvider cm01.contoso.com
```

### Example 2

Export collected data in folder ".\2019-2-6\cm01.contoso.com" to HTML report...

```powershell
# basic example

Export-CMHealthCheckHTML -ReportFolder "2019-2-6\cm01.contoso.com" -OutputFolder "c:\reports" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Verbose

# fancy example

Export-CMHealthCheckHTML -ReportFolder "2019-2-6\cm01.contoso.com" -OutputFolder "c:\reports" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Theme 'Ocean' -DynamicTableRows -Verbose
```

### Example 3

Export collected data in folder ".\2019-2-6\cm01.contoso.com" to Microsoft Word report...

```powershell

Export-CMHealthCheck -ReportFolder "2019-2-6\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose

```
