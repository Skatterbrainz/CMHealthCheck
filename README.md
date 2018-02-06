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

## Notes

* Refer to the markdown files under \Docs for syntax and examples.
* Tested with the following environments:
   * ConfigMgr 2012 R2, 1610, 1702, 1706, 1709, 1710, 1711
   * Windows Server 2012, 2012 R2, 2016
   * SQL Server 2012, 2016 SP1
   * Windows 10 1703, 1709
   * Office 2013, 2016
   * PowerShell 3.0, 5.0, 5.1
  
* If you like it, please share with others.  If you hate it, tell me why so I can improve on it?
* Please submit bugs, comments, requests via the "Issues" link above.

## Removal and Cleanup

* To remove CMHealthCheck module and related files, use the Remove-Module or Uninstall-Module cmdlets.
