# CMHealthCheck
ConfigMgr Health Check Reporting PowerShell functions

## Summary

CMHealthCheck is intended to make the former standalone PowerShell scripts "get-cm-inventory.ps1" and "export-cm-healthcheck.ps1" easier to invoke.  The scripts have been refactored into a simple PowerShell module, with two public functions: Get-CMHealthCheck and Export-CMHealthCheck.  The required support XML data files are now invoked via URI from GitHub gist locations (on this account), but they can be downloaded for offline use as well.

## Installation

_Part 1 - Data Collection_

* Log onto SCCM site server with full admin credentials (SCCM, Windows and SQL Server instance)
* Open PowerShell console using Run as Administrator
* Enter: Install-Module CMHealthCheck
* Enter: Get-CMHealthCheck -SmsProvider "hostname" (add -Verbose for more detailed output) and press Enter
* Collect output files from $env:USERPROFILE\Documents\YYYY-MM-DD\hostname
* Copy to machine which has Office 2013 or 2016 installed (part 2)

_Part 2 - Reporting_

* Temp
* Temp

## Removal and Cleanup

To remove CMHealthCheck module and related files, use the Remove-Module cmdlet.
