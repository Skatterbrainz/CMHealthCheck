# CMHealthCheck
ConfigMgr Health Check Reporting PowerShell functions

## Summary

CMHealthCheck is intended to make the former standalone PowerShell scripts "get-cm-inventory.ps1" and "export-cm-healthcheck.ps1" easier to invoke.  The scripts have been refactored into a simple PowerShell module, with two public functions: Get-CMHealthCheck and Export-CMHealthCheck.  The required support XML data files are now invoked via URI from GitHub gist locations (on this account), but they can be downloaded for offline use as well.

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

## Syntax: Get-CMHealthCheck

* Get-CMHealthCheck -SmsProvider "cm01.contoso.com" ...

### Parameters
* SmsProvider

   FQDN of ConfigMgr site server.  Example: "cm01.contoso.com"
   
* OutputFolder 

   optional path for output files. Default is $env:USERPROFILE\Documents

* NumberOfDays 

   optional age of status logs to inspect. Default is 7

* HealthcheckFilename 
  
  optional path to cmhealthcheck.xml. Default is Git Gist URL (see Notes)

* Overwrite 
  
  optional switch to force replacing output if on same date

* NoHotFix 
  
  optional switch to skip auditing of installed hotfixes / may save time

* Verbose (ummm, yeah)

## Syntax: Export-CMHealthCheck

* Export-CMHealthCheck -ReportFolder "path to output files" ...
* -Detailed 
  * optional switch to force more verbose reporting output / strongly recommended!
* -CoverPage 
  * optional Windows theme cover page. Default is "Slices (light)"
* -CustomerName (optional name of SCCM site server owner. Default is "Customer Name")
* -AuthorName
* -CopyrightName
* -HealthcheckFilename
* -MessagesFilename
* -HealthcheckDebug
* -Overwrite 
  * (ignore this, I had no sleep and a cat that wouldn't leave me alone)

## Notes

## Removal and Cleanup

To remove CMHealthCheck module and related files, use the Remove-Module cmdlet.
