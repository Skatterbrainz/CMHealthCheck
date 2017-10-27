# CMHealthCheck

ConfigMgr Health-Check Reporting PowerShell functions

## Summary

CMHealthCheck is intended to make the former standalone PowerShell scripts "get-cm-inventory.ps1" and "export-cm-healthcheck.ps1" easier to invoke as a module from PowerShell Gallery.  Both scripts have been refactored into a simple PowerShell module, with two public functions: Get-CMHealthCheck and Export-CMHealthCheck.  The required support XML data files are now invoked via URI from GitHub gist locations (on this account), but they can be downloaded for offline use as well.

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

   ```powershell
   Get-CMHealthCheck -SmsProvider "cm01.contoso.com" ...
   ```

### Parameters

* **SmsProvider**

   FQDN of ConfigMgr site server.  Example: "cm01.contoso.com"
   
* **OutputFolder**

   _Optional_ Path for output files. Default is $env:USERPROFILE\Documents. The script will create two (2) folders 
   beneath this location: _Logs, and another using YYYY-MM-DD, with a sub-folder named after the SmsProvider value.

* **NumberOfDays**

   _Optional_ Integer value, number of days to go back for status logs to inspect. Default is 7

* **HealthcheckFilename**
  
  _Optional_ Path or URI to cmhealthcheck.xml, which provides rules for gathering data. Default is Git Gist URL <https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml>

* **Overwrite**
  
  _Optional_ Switch parameter to force replacing output if on same date.  If the function has been executed on a given ConfigMgr site server on the same date, there will already be a "YYYY-MM-DD\hostname" output file with data files.  Without the -Overwrite switch, the default behavior is to display a warning and abort.

* **NoHotFix**
  
  _Optional_ Switch parameter to skip auditing of installed hotfixes.  This may save time when re-running a data collection in test environments.

* **Verbose** 

   (ummm, yeah)

### Examples

   ```powershell
   Get-CMHealthCheck -SmsProvider "cm01.contoso.com" -OutputFolder "C:\Temp" -NumberOfDays 30 
   ```
   
## Syntax: Export-CMHealthCheck

   ```powershell
   Export-CMHealthCheck -ReportFolder "path to output files" ...
   ```
   
### Parameters

* **ReportFolder**

   _Mandatory_ Path to where the collected data files reside from using Get-CMHealthCheck. This can be a local path or a remote UNC path.
   
* **Detailed**

   _Optional_ Switch parameter to force more verbose reporting output / strongly recommended!
   
* **CoverPage**

   _Optional_ Name of Office cover page.  List of valid names varies based on the version of Office installed.
   
   Default is "Slice (Light)"
   
   Word 2016 names: Austin, Banded, Facet, Filigree, Grid, Integral, Ion (Dark), Ion (Light), Motion, Retrospect, Semaphore, Sideline, Slice (Dark), Slice (Light), Viewmaster, Whisp
   
* **CustomerName**

   _Optional_ Name of customer or organization who owns the ConfigMgr site server being audited. Default is "Customer Name"

* **AuthorName**

   _Optional_ Name of author generating the report (you?).  Default is "Your Name"
   
* **CopyrightName**

   _Optional_ Name to place in footer of every page along with (C)YYYY .....  Default is "Your Company Name"
   
* **HealthcheckFilename**

   _Optional_ Path or URI to cmhealthcheck.xml, which provides rules for gathering data. Default is Git Gist URL <https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml>
   
* **MessagesFilename**

  _Optional_ Path or URI to messages.xml, which provides status value message lookups. Default is Git Gist URL <https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/Messages.xml>
  
* **HealthcheckDebug**

   _Optional_ Switch parameter to enable additional verbose output
   
* **Overwrite**
   
   ignore this, I had no sleep and a cat that wouldn't leave me alone

### Examples

   ```powershell
   Export-CMHealthCheck -ReportFolder "C:\Temp\2017-10-23\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "Mike Hunt" -CopyrightName "Fubar LLC"
   ```
   
   ```powershell
   Export-CMHealthCheck -ReportFolder "C:\Temp\2017-10-23\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "Mike Hunt" -CopyrightName "Fubar LLC" -HealthcheckDebug -Overwrite
   ```

   ```powershell
   Export-CMHealthCheck -ReportFolder "C:\Temp\2017-10-23\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "Mike Hunt" -CopyrightName "Fubar LLC" -CoverPage "Ion (Dark)" -HealthcheckFilename "C:\Temp\cmhealthcheck.xml" -MessagesFile "C:\Temp\messages.xml" -HealthcheckDebug -Overwrite
   ```

## Notes

* If you like it, please share with others.  If you hate it, tell me why so I can improve on it?
* Please submit bugs, comments, requests via the "Issues" link above.

## Removal and Cleanup

To remove CMHealthCheck module and related files, use the Remove-Module cmdlet.
