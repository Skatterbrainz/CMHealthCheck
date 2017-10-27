#requires -RunAsAdministrator
#requires -version 3
<#
.SYNOPSIS
    Get-CMHealthCheck collects SCCM hierarchy and site server data

.DESCRIPTION
    Get-CMHealthCheck collects SCCM hierarchy and site server data
    and stores the information in multiple XML data files which are then
    processed using the Export-CM-Healthcheck.ps1 script to render
    a final MS Word report.

.PARAMETER ReportFolder
    [string] [required] Path to output data folder

.PARAMETER SmsProvider
    [string] [required] FQDN of SCCM site server

.PARAMETER NumberOfDays
    [int] [optional] Number of days to go back for alerts in logs
    default = 7

.PARAMETER HealthcheckFilename
    [string] [optional] Name of configuration file
    default is cmhealthcheck.xml

.PARAMETER Overwrite
    [switch] [optional] Overwrite existing output folder if found
    Folder is named by datestamp, so this only applies when
    running repeatedly on the same date

.PARAMETER NoHotfix
    [switch] [optional] Suppress hotfix inventory
    Can save significant runtime

.NOTES
	See GitHub Wiki for version updates and details

	Thanks to:
    Base script (the hardest part) created by Rafael Perez (www.rflsystems.co.uk)
    Word functions copied from Carl Webster (www.carlwebster.com)
    Word functions copied from David O'Brien (www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/)

    NOTE: This script was tested on SCCM from 2012 R2 up to 1703 Primary and CAS hierarchy environments

    Support: Database name must be CM_<SITECODE> (you need to adapt the queries if not this format)

    Security Rights: user running this tool should have the following rights:
        SQL Server (serveradmin) to be able to see database / cpu stats
        SCCM Database (db_owner) used to create/drop user-defined functions
        msdb Database (db_datareader) used to read backup information
        read-only analyst on the SCCM console
        local administrator on all computer (used to remotely connect to the registry and services)
        firewall allowing 1433 (or equivalent) to all SQL Servers (including SQL Express on Secondary Site)
        Remote WMI/DCOM firewall - http://msdn.microsoft.com/en-us/library/jj980508(v=winembedded.81).aspx
        Remote WUA - http://msdn.microsoft.com/en-us/library/windows/desktop/aa387288%28v=VS.85%29.aspx

    Comments: To get the free disk space, enable the Free Space (MB) for the Logical Disk

.EXAMPLE
    .\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NumberofDays:30
	.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -Overwrite -Verbose
	.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -HealthcheckDebug -Verbose
	.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NoHotfix
	.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -OutputFolder "c:\temp"
#>

function Get-CMHealthCheck {
	[CmdletBinding(ConfirmImpact="Low")]
	param (
		[Parameter(
			Mandatory = $True, 
			HelpMessage = "Enter the SMS Provider computer name",
			ValueFromPipeline=$True
		)] 
			[ValidateNotNullOrEmpty()]
			[string] $SmsProvider,
		[parameter(Mandatory=$False, HelpMessage="Path for output data files")]
			[ValidateNotNullOrEmpty()]
			[string] $OutputFolder = "$($env:USERPROFILE)\Documents",
		[Parameter(Mandatory = $False, HelpMessage = "Number of Days for HealthCheck")] 
			[int] $NumberofDays = 7,
		[Parameter (Mandatory = $False, HelpMessage = "HealthCheck query file name")] 
			[string] $Healthcheckfilename = 'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml',
		[Parameter(Mandatory = $False, HelpMessage = "Overwrite existing report?")] 
			[switch] $Overwrite,
		[Parameter(Mandatory=$False, HelpMessage="Skip hotfix inventory")]
			[switch] $NoHotfix
	)

	Start-Transcript -Path ".\Get-CM-Inventory-Runtime.log"

	$ScriptVersion = "1710.01"
	$startTime     = Get-Date
	$currentFolder = $PWD.Path
	if ($currentFolder.substring($currentFolder.Length-1) -ne '\') { $currentFolder+= '\' }
	$logFolder     = $OutputFolder + '\_Logs\'
	$reportFolder  = $OutputFolder + '\' + (Get-Date -UFormat "%Y-%m-%d") + '\' + $SmsProvider + '\'
	$component     = ($MyInvocation.MyCommand.Name -replace '.ps1', '')
	$logfile       = $logFolder + $component + ".log"
	$poshversion   = $PSVersionTable.PSVersion.Major
	$osversion     = (Get-WmiObject -Class Win32_OperatingSystem).Caption
	$FormatEnumerationLimit = -1
	$Error.Clear()
	$bLogValidation = $False

	Write-Host "Get-CMHealthCheck - version $ScriptVersion"
	Write-Host "Gathering site and server information"

	Write-Verbose "-------------------------------------"
    Write-Verbose "Report Folder: $reportFolder"
	Write-Verbose "Running Powershell version: $poshversion"
	if (!(Test-Powershell64bit)) {
		Write-Error "Powershell is not 64bit, no futher action taken"
		break
	}
	Write-Verbose "Running Powershell 64 bits"
	Write-Verbose "Windows Version: $osversion"
	Write-Verbose "SMS Provider: $smsprovider"

	try {
		if (-not (Test-Admin)) {
			Write-Host "You are not running PowerShell as Administrator (run as Administrator), no futher action taken" -ForegroundColor Red
			break		
		}

		if (Test-Path -Path $reportFolder) {
			if ($Overwrite -eq $true) {
				Write-Verbose "removing previous output folder $($reportFolder)..."
				Remove-Item -Path "$($reportFolder)" -Recurse -Force
			}
			else {
				Write-Host "Folder $reportFolder already exist, no futher action taken" -ForegroundColor Red
				break
			}
		}

		if (Test-Folder -Path $logFolder) {
			try {
				New-Item ($logFolder + 'Test.log') -Type File -Force | Out-Null 
				Remove-Item ($logFolder + 'Test.log') -Force | Out-Null 
			}
			catch {
				Write-Error "Unable to read/write file on $logFolder folder, no futher action taken"
				break
			}
		}
		else {
			Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
			break
		}
		$bLogValidation = $true

		Write-Verbose "--------------- importing cmhealthcheck.xml ---------------"
		if ($Healthcheckfilename.StartsWith('http')) {
			Write-Verbose "importing xml from remote URI: $healthcheckfilename"
			try {
				[xml]$HealthCheckXML = Invoke-RestMethod -Uri $Healthcheckfilename
			}
			catch {
				Write-Error "Failed to import data from Uri: $HealthcheckFilename"
				Write-Warning "If no Internet access is allowed, use -HealthcheckFilename '.\cmhealthcheck.xml'"
				break
			}
			Write-Verbose "configuration XML data loaded successfully"
		}
		else {
			if (!(Test-Path -Path ($currentFolder + $healthcheckfilename))) {
				Write-Error "File $($currentFolder)$($healthcheckfilename) does not exist, no futher action taken"
				break
			}
			else { 
				[xml]$HealthCheckXML = Get-Content ($currentFolder + $healthcheckfilename) 
			}
		}

		if (Test-Folder -Path $reportFolder) {
			try {
				New-Item ($reportFolder + 'Test.log') -Type file -Force | Out-Null 
				Remove-Item ($reportFolder + 'Test.log') -Force | Out-Null 
			}
			catch {
				Write-Host "Unable to read/write file on $reportFolder folder, no futher action taken" -ForegroundColor Red
				break
			}
		}
		else {
			Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
			break
		}
		
		if (($Overwrite) -and (Test-Path $logfile)) {
			Remove-Item $logfile -Force
			Write-Verbose "previous log file cleared via overwrite request"
		}
		Write-Verbose "-------------- connecting to site ---------------------"
	
		$WMISMSProvider = Get-CmWmiObject -Class "SMS_ProviderLocation" -NameSpace "Root\SMS" -ComputerName $smsprovider -LogFile $logfile
		$SiteCodeNamespace = $WMISMSProvider.SiteCode
		Write-Verbose "Site Code: $SiteCodeNamespace"
		
		$WMISMSSite = Get-CmWmiObject -Class "SMS_Site" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -Filter "SiteCode = '$SiteCodeNamespace'" -ComputerName $smsprovider -LogFile $logfile
		$SMSSiteServer = $WMISMSSite.ServerName
		Write-Verbose "Site Server: $($WMISMSSite.ServerName)"
		Write-Verbose "Site Version: $($WMISMSSite.Version)" 

		if (-not ($WMISMSSite.Version -like "5.*")) {
			Write-Verbose "SCCM Site $($WMISMSSite.Version) not supported. No further action taken"
			break
		}
		
		$SQLServerName  = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Server'
		$SQLServiceName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server' -KeyValue 'Service Name'
		$SQLPort   = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Port'
		$SQLDBName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Database Name'
		
		# parse when finding default instance vs named instance
		if ($SQLDBName.IndexOf('\') -ge 0) {
			$SQLDBName = $SQLDBName.Split("\")[1]
		}
		
		Write-Verbose "--------------- getting sql server info --------------------"	
		Write-Verbose "SQLServerName: $SQLServerName"
		Write-Verbose "SQLServiceName: $SQLServiceName"
		Write-Verbose "SQLPort: $SQLPort"
		Write-Verbose "SQLDBName: $SQLDBName"

		$arrServers = @()
		$WMIServers = Get-CmWmiObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where NetworkOSPath not like '%.microsoft.com' and Type in (1,2,4,8)" -ComputerName $smsprovider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
		foreach ($WMIServer in $WMIServers) { 
			$arrServers += $WMIServer.NetworkOSPath -replace '\\', '' 
		}
		if ($arrServers.Count -gt 0) {
			Write-Verbose $("Servers discovered: " + $arrServers -join(", "))
		}
		else {
			Write-Verbose "no servers discovered."
		}
		Write-Verbose "----------------- creating temp data table ------------------"
		$Fields = @("TableName", "XMLFile")
		$ReportTable = New-CmDataTable -TableName $tableName -Fields $Fields

		$Fields = @("SiteServer", "SQLServer","DBName","SiteCode","NumberOfDays","HealthCheckFileName")
		$ConfigTable = New-CmDataTable -TableName $tableName -Fields $Fields

		$row = $ConfigTable.NewRow()
		$row.SiteServer   = $SMSSiteServer
		$row.SQLServer    = $SQLServerName
		$row.DBName       = $SQLDBName
		$row.SiteCode     = $SiteCodeNamespace
		$row.NumberOfDays = [System.Convert]::ToInt32($NumberOfDays)
		$row.HealthCheckFileName = $HealthCheckFileName

		$ConfigTable.Rows.Add($row)
		Write-Verbose "Exporting XML to $($reportFolder)config.xml"
		, $ConfigTable | Export-Clixml -Path ($reportFolder + 'config.xml')

		$sqlConn = Get-SQLServerConnection -SQLServer "$SQLServerName,$SQLPort" -DBName $SQLDBName
		$sqlConn.Open()

		Write-Verbose "SQL Query: Creating Functions"
		$functionsSQLQuery = @"
CREATE FUNCTION [fn_CM12R2HealthCheck_ScheduleToMinutes](@Input varchar(16))
RETURNS bigint
AS
BEGIN
	if (ISNULL(@Input, '') <> '')
	begin
		declare @hex varchar(64), @flag char(3), @minute char(6), @hour char(5), @day char(5), @Cnt tinyint, @Len tinyint, @Output bigint, @Output2 bigint = 0
		
		set @hex = @Input

		SET @HEX=REPLACE (@HEX,'0','0000')
		set @hex=replace (@hex,'1','0001')
		set @hex=replace (@hex,'2','0010')
		set @hex=replace (@hex,'3','0011')
		set @hex=replace (@hex,'4','0100')
		set @hex=replace (@hex,'5','0101')
		set @hex=replace (@hex,'6','0110')
		set @hex=replace (@hex,'7','0111')
		set @hex=replace (@hex,'8','1000')
		set @hex=replace (@hex,'9','1001')
		set @hex=replace (@hex,'A','1010')
		set @hex=replace (@hex,'B','1011')
		set @hex=replace (@hex,'C','1100')
		set @hex=replace (@hex,'D','1101')
		set @hex=replace (@hex,'E','1110')
		set @hex=replace (@hex,'F','1111')
		
		select @Flag = SUBSTRING(@hex,43,3), @minute = SUBSTRING(@hex,46,6), @hour = SUBSTRING(@hex,52,5), @day = SUBSTRING(@hex,57,5)

		if (@flag = '010') --SCHED_TOKEN_RECUR_INTERVAL
		BEGIN
			set @Cnt = 1
			set @Len = LEN(@minute)
			set @Output = CAST(SUBSTRING(@minute, @Len, 1) AS bigint)
			
			WHILE(@Cnt < @Len) BEGIN
			SET @Output = @Output + POWER(CAST(SUBSTRING(@minute, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			SET @Cnt = @Cnt + 1
			END
			set @Output2 = @Output
			
			set @Cnt = 1
			set @Len = LEN(@hour)
			set @Output = CAST(SUBSTRING(@hour, @Len, 1) AS bigint)
			
			WHILE(@Cnt < @Len) BEGIN
			SET @Output = @Output + POWER(CAST(SUBSTRING(@hour, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			SET @Cnt = @Cnt + 1
			END		
			set @Output2 = @Output2 + (@Output*60)
			
			set @Cnt = 1
			set @Len = LEN(@day)
			set @Output = CAST(SUBSTRING(@day, @Len, 1) AS bigint)
			
			WHILE(@Cnt < @Len) BEGIN
			SET @Output = @Output + POWER(CAST(SUBSTRING(@day, @Len - @Cnt, 1) * 2 AS bigint), @Cnt)
			SET @Cnt = @Cnt + 1
			END		
			set @Output2 = @Output2 + (@Output*60*24)
		END
		ELSE
			set @Output2 = -1
	end
	else
		set @Output2 = -2
		
	return @Output2
END
"@
		$SqlCommand = $sqlConn.CreateCommand()
		$SqlCommand.CommandTimeOut = 0
		$SqlCommand.CommandText = $functionsSQLQuery
		try {
			$SqlCommand.ExecuteNonQuery() | Out-Null
		}
		catch {}
		$SqlCommand = $null

		$arrSites = @()
		$SqlCommand = $sqlConn.CreateCommand()

		$executionquery = "select distinct st.SiteCode, (select top 1 srl2.ServerName from v_SystemResourceList srl2 where srl2.RoleName = 'SMS Provider' and srl2.SiteCode = st.SiteCode) as ServerName from v_Site st"
		Write-Verbose "SQL Query...`n$executionquery"

		$SqlCommand.CommandTimeOut = 0
		$SqlCommand.CommandText = $executionquery

		Write-Verbose "--------------- querying database --------------------"	
		Write-Verbose "processing query to sql data adapter..."
		$DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
		$dataset     = New-Object System.Data.Dataset
		try {
			$DataAdapter.Fill($dataset) | Out-Null
		}
		catch {
			Write-Error "oh shit! you done opened a can o whoopass!"
		}
		Write-Verbose "data adapter is good!"
		foreach($row in $dataset.Tables[0].Rows) { 
			$arrSites += "$($row.SiteCode)@$($row.ServerName)" 
		}
		Write-Verbose $("Sites discovered: " + $arrSites -join(", "))

		$SqlCommand = $null

		##section 1 = information that needs be collected against each site
		Write-Host "Phase 1 of 6"
		
		foreach ($Site in $arrSites) {
			$arrSiteInfo = $Site.Split("@")
			$PortInformation = Get-CmWmiObject -query "select * from SMS_SCI_Component where FileType=2 and ItemName='SMS_MP_CONTROL_MANAGER|SMS Management Point' and ItemType='Component' and SiteCode='$($arrSiteInfo[0])'" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -ComputerName $smsprovider -LogFile $logfile
			foreach ($portinfo in $PortInformation) {
				$HTTPport  = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISPortsList"}).Value1
				$HTTPSport = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISSSLPortsList"}).Value1
			}
			ReportSection -HealthCheckXML $HealthCheckXML -Section '1' -sqlConn $sqlConn -SiteCode $arrSiteInfo[0] -NumberOfDays $NumberOfDays -ServerName $arrSiteInfo[1] -ReportTable $ReportTable -LogFile $logfile 
		} # foreach

		##section 2 = information that needs be collected against each computer. should not be site specific. query will run only against the higher site in the hierarchy
		Write-Host "Phase 2 of 6"
		
		foreach ($server in $arrServers) { 
			ReportSection -HealthCheckXML $HealthCheckXML -Section '2' -sqlConn $sqlConn -siteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $server -ReportTable $ReportTable -LogFile $logfile 
		}
		
		##section 3 = database analisys information, running on all sql servers in the hierarchy. should not be site specific as it connects to the "master" database
		Write-Host "Phase 3 of 6"
		
		$DBServers = Get-CmWmiObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where RoleName = 'SMS SQL Server'" -ComputerName $smsprovider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
		foreach ($DB in $DBServers) { 
			$DBServerName = $DB.NetworkOSPath.Replace('\',"") 
			Write-Log -Message ("Analysing SQLServer: $DBServerName") -LogFile $logfile
			if ($SQLServerName.ToLower() -eq $DBServerName.ToLower()) { 
				$tmpConnection = $sqlConn 
			}
			else {
				$tmpConnection = Get-SQLServerConnection -SQLServer "$DBServerName,$SQLPort" -DBName "master"
				$tmpConnection.Open()
			}
			try {
				ReportSection -HealthCheckXML $HealthCheckXML -Section '3' -sqlConn $tmpConnection -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $DBServerName -ReportTable $ReportTable -LogFile $logfile
			}
			finally {
				if ($SQLServerName.ToLower() -ne $DBServerName.ToLower()) { $tmpConnection.Close()  }
			}
		} # foreach

		##Section 4 = Database analysis against whole SCCM infrastructure, query will run only against top SQL Server
		Write-Host "Phase 4 of 6"
		
		ReportSection -HealthCheckXML $HealthCheckXML -Section '4' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile

		##Section 5a = summary information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
		Write-Host "Phase 5 of 6"
		
		ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile
		
		##Section 5b = detailed information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
		Write-Verbose "**** ENTERING SECTION 5b ****"
		
		ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -Detailed -LogFile $logfile

		##Section 6 = troubleshooting information
		Write-Host "Phase 6 of 6"
		
		ReportSection -HealthCheckXML $HealthCheckXML -Section '6' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile
	}
	catch {
		Write-Verbose "ERROR/EXCEPTION: general unhandled exception"
		Write-Verbose "The following error occurred, no futher action taken"
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		Write-Verbose "Error $errorCode : $errorMessage" 
		Write-Verbose "Full Error Message Error $($error[0].ToString())"
		$Error.Clear()
	}
	finally {
		#close sql connection
		Write-Host "Finishing up"
		Write-Verbose "Closing SQL connection"
		if ($sqlConn -ne $null) {
			Write-Verbose "SQL Query: Deleting Functions"
			$functionsSQLQuery = @"
IF OBJECT_ID (N'fn_CM12R2HealthCheck_ScheduleToMinutes', N'FN') IS NOT NULL
	DROP FUNCTION fn_CM12R2HealthCheck_ScheduleToMinutes;
"@
			try {
				$SqlCommand = $sqlConn.CreateCommand()
				$SqlCommand.CommandTimeOut = 0
				$SqlCommand.CommandText = $functionsSQLQuery
				try {
					$SqlCommand.ExecuteNonQuery() | Out-Null 
				}
					catch {}
				$SqlCommand = $null
				$sqlConn.Close() 
			}
			catch {}
		}
		if ($ReportTable -ne $null) { , $ReportTable | Export-CliXml -Path ($reportFolder + 'report.xml') }

		if ($bLogValidation -eq $false) {
			Write-Host "Ending HealthCheck CollectData"
			Write-Verbose "log validation not enabled"
		}
		else {
			Write-Verbose "Ending HealthCheck CollectData"
		}
	}
	$StopTime = Get-Date
	$RunSecs  = ((New-TimeSpan -Start $StartTime -End $StopTime).TotalSeconds).ToString()
	$ts       = [timespan]::FromSeconds($RunSecs)
	$RunTime  = $ts.ToString("hh\:mm\:ss")
	Write-Output "Processing completed. Total runtime: $RunTime (hh`:mm`:ss)"
	try { 
		Stop-Transcript -ErrorAction SilentlyContinue
	}
	catch {}
}

Export-ModuleMember -Function Get-CMHealthCheck
