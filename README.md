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

## Notes

* Tested with the following environments:
   * ConfigMgr 2012 R2, 1610 to 1902.2 (GA and Tech Preview)
   * Windows Server 2012, 2012 R2, 2016
   * SQL Server, 2012, 2014, 2016, 2017
   * Windows 10, 1703 to 1809
   * Office 2013, 2016 / 365
   * PowerShell 5.1 (previous versions no longer supported. Time to move up!)
  
* If you like it, please share with others.  If you hate it, tell me why so I can improve on it?
* Please submit bugs, comments, requests via the "Issues" link above.

## Removal and Cleanup

* To remove CMHealthCheck module and related files, use the Remove-Module or Uninstall-Module cmdlets.

## Examples

* Refer to the "Docs" folder for markdown documents for each function.