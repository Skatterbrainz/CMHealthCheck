# CMHealthCheck

ConfigMgr Health-Check Reporting PowerShell functions

## Summary

CMHealthCheck is intended to make it easier to generate health check reports for Configuration Manager
sites.  The reports also include general configuration information.  The process is divided into two (2)
parts: Data Collection and Report Generation.

## Installation and Usage

```powershell
Install-Module CMHealthCheck
```

Use ```Get-Command -Module CMHealthCheck``` and ```Get-Help <function>``` to view latest examples and help.

Commands / Functions:

* [Get-CMHealthCheck](/Docs/Get-CMHealthCheck.md)
* [Get-CMHealthCheckSummary](/Docs/Get-CMHealthCheckSummary.md)
* [Export-CMHealthReport](/Docs/Export-CMHealthReport.md)
* [Invoke-CMHealthCheck](/Docs/Invoke-CMHealthCheck.md)

## Notes

* Tested with the following environments:
   * ConfigMgr 2012 R2 to the latest Technical Preview Current Branch
   * Windows Server 2012, 2012 R2, 2016, 2019, 2022
   * SQL Server, 2012, 2014, 2016, 2017, 2019, 2022
   * Windows 10, 1703 to 21H1, Windows 11 (Preview)
   * Office 2013, 2016 and 2019 / 365
   * PowerShell 5.1 (previous versions no longer supported. Time to move up!)
  
* If you like it, please share with others.  If you hate it, tell me why so I can improve on it?
* Please submit bugs, comments, requests via the "Issues" link above.

## Removal and Cleanup

* To remove CMHealthCheck module and related files, use the Remove-Module or Uninstall-Module cmdlets.

## Examples

* Samples are provided in the Samples folder within the module installation path
* Refer to the "Docs" folder for markdown documents for each function.
