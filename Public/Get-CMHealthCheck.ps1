function Get-CMHealthCheck {
	<#
	.SYNOPSIS
		Extract ConfigMgr Site data
	.DESCRIPTION
		Exracts SCCM hierarchy and site server data
		and stores the information in multiple XML data files which are then
		processed using the Export-CM-Healthcheck.ps1 script to render
		a final MS Word report.
	.PARAMETER OutputFolder
		Path to output data folder, default is "My Documents"
	.PARAMETER SmsProvider
		FQDN of SCCM site server
	.PARAMETER NumberOfDays
		Number of days to go back for alerts in logs (default = 7)
	.PARAMETER HealthcheckFilename
		Name of configuration file (default is .\assets\cmhealthcheck.xml)
	.PARAMETER Overwrite
		Overwrite existing output folder if found.
		Folder is named by datestamp, so this only applies when
		running repeatedly on the same date
	.PARAMETER NoHotfix
		Suppress hotfix inventory. Can save significant runtime
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NumberofDays:30
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -Overwrite -Verbose
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -HealthcheckDebug -Verbose
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -NoHotfix
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -OutputFolder "c:\temp"
	.EXAMPLE
		.\Get-CMHealthCheck -SmsProvider cm01.contoso.com -OutputFolder "c:\temp" -HealthcheckFilename ".\healthcheck.xml"
	.NOTES
		* Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
		* Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
		* Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
		* Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference
		* Support: Database name must be CM_<SITECODE> (you need to adapt the queries if not this format)

		* Security Rights: user running this tool should have the following rights:
		  - SQL Server (serveradmin) to be able to see database / cpu stats
		  - SCCM Database (db_owner) used to create/drop user-defined functions
		  - msdb Database (db_datareader) used to read backup information
		  - read-only analyst on the SCCM console
		  - local administrator on all computer (used to remotely connect to the registry and services)
		  - firewall allowing 1433 (or equivalent) to all SQL Servers (including SQL Express on Secondary Site)
		  - Remote WMI/DCOM firewall - http://msdn.microsoft.com/en-us/library/jj980508(v=winembedded.81).aspx
		  - Remote WUA - http://msdn.microsoft.com/en-us/library/windows/desktop/aa387288%28v=VS.85%29.aspx
		  - Comments: To get the free disk space, enable the Free Space (MB) for the Logical Disk
	.LINK
		https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Get-CMHealthCheck.md
	#>
	[CmdletBinding(ConfirmImpact="Low")]
	param (
		[parameter()] [ValidateNotNullOrEmpty()][string] $SmsProvider = "$(($env:COMPUTERNAME, $env:USERDNSDOMAIN) -join '.')",
		[parameter()][ValidateNotNullOrEmpty()][string] $OutputFolder = "$([System.Environment]::GetFolderPath('Personal'))",
		[parameter()][ValidateRange(1,365)][int] $NumberOfDays = 7,
		[parameter()][string] $Healthcheckfilename = "",
		[parameter()][switch] $OverWrite,
		[parameter()][switch] $NoHotfix
	)

	$startTime      = Get-Date
	$currentFolder  = $(Get-Location).Path
	if ($currentFolder.substring($currentFolder.Length-1) -ne '\') { $currentFolder+= '\' }
	$logFolder      = "$OutputFolder\_Logs\"
	$reportFolder   = "$OutputFolder\$(Get-Date -UFormat "%Y-%m-%d")\$SmsProvider\"
	if (-not(Test-Path $logFolder)) {
		Write-Log -Message "creating log folder: $logFolder"
		mkdir -Path $logFolder -Force
	}
	if (-not(Test-Path $reportFolder)) {
		Write-Log -Message "creating report folder: $reportFolder"
		mkdir -Path $reportFolder -Force
	}
	$logfile        = Join-Path -Path $logFolder -ChildPath "Get-CMHealthCheck.log"
	$poshversion    = $PSVersionTable.PSVersion.Major
	$osversion      = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
	$Error.Clear()
	$bLogValidation = $False
	$ModuleVer      = $(Get-Module "CMHealthCheck").Version -join '.'
	$ModulePath     = Split-Path (Get-Module "CMHealthCheck").Path -Parent
	if ([string]::IsNullOrEmpty($Healthcheckfilename)) {
		$Healthcheckfilename = "$ModulePath\assets\cmhealthcheck.xml"
		Write-Log -Message "Healthcheck file....: $HealthCheckFileName" -LogFile $logfile
	}
	Write-Host "CMHealthCheck $ModuleVer"
	Write-Host "Gathering site and server information"

	Write-Log -Message "----------------- BEGIN PROCESSING --------------------" -LogFile $logfile
	Write-Log -Message "Module version......: $ModuleVer" -LogFile $logfile
	Write-Log -Message "Report Folder.......: $reportFolder" -LogFile $logfile
	Write-Log -Message "Powershell version..: $poshversion" -LogFile $logfile
	if (!(Test-Powershell64bit)) {
		Write-Error "Powershell is not 64bit, yo G, we outta here."
		break
	}
	Write-Log -Message "PowerShell mode.....: 64-bit" -LogFile $logfile
	Write-Log -Message "Windows Version.....: $osversion" -LogFile $logfile
	Write-Log -Message "SMS Provider........: $smsprovider" -LogFile $logfile

	try {
		if (-not (Test-Admin)) {
			Write-Host "You are not running PowerShell as Administrator (run as Administrator), no futher action taken" -ForegroundColor Red
			break
		}
		if (Test-Path -Path $reportFolder) {
			if ($Overwrite -eq $true) {
				Write-Log -Message "removing previous output folder $($reportFolder)..." -LogFile $logfile
				Remove-Item -Path "$($reportFolder)" -Recurse -Force
			}
			else {
				Write-Host "Folder $reportFolder already exists. Use -Overwrite to replace." -ForegroundColor Red
				break
			}
		}

		$bLogValidation = $true

		Write-Log -Message "--------------- importing cmhealthcheck.xml ---------------" -LogFile $logfile
		[xml]$HealthCheckXML = Get-CmHealthCheckFile -XmlSource $HealthCheckFilename
		if (!(Test-Folder -Path $reportFolder)) {
			Write-Log -Message "Unable to create $reportFolder" -Severity 3 -LogFile $logfile
			break
		}
		if (($Overwrite) -and (Test-Path $logfile)) {
			Remove-Item $logfile -Force
			Write-Log -Message "previous log file cleared via overwrite request" -LogFile $logfile
		}
		Write-Log -Message "-------------- connecting to site ---------------------"

		$WMISMSProvider = Get-CmWmiObject -Class "SMS_ProviderLocation" -NameSpace "Root\SMS" -ComputerName $smsprovider -LogFile $logfile
		$SiteCodeNamespace = $($WMISMSProvider.SiteCode -join '').SubString(0,3) # thanks to Chris Shilt (07/21/20)
		if ([string]::IsNullOrEmpty($SiteCodeNameSpace)) {
			Write-Host "Error: Unable to connect to $SmsProvider. Exit." -ForegroundColor Red
			Write-Log "unable to connect to $SmsProvider. Exiting here." -Severity 3 -LogFile $logfile
			break
		}
		Write-Log -Message "Site Code........: $SiteCodeNamespace" -LogFile $logfile

		$WMISMSSite = Get-CmWmiObject -Class "SMS_Site" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -Filter "SiteCode = '$SiteCodeNamespace'" -ComputerName $smsprovider -LogFile $logfile
		$SMSSiteServer  = $WMISMSSite.ServerName
		$SMSSiteVersion = $WMISMSSite.Version
		Write-Log -Message "Site Server......: $SMSSiteServer" -LogFile $logfile
		Write-Log -Message "Site Version.....: $SMSSiteVersion" -LogFile $logfile

		if ($WMISMSSite.Version -lt 5) {
			Write-Log -Message "ConfigMgr Site $SMSSiteVersion not supported. No further action taken" -Severity 3 -LogFile $logfile
			break
		}

		$SQLServerName  = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Server'
		$SQLServiceName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server' -KeyValue 'Service Name'
		$SQLPort        = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Port'
		$SQLDBName      = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile $logfile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Database Name'

		# parse when finding default instance vs named instance
		if ($SQLDBName.IndexOf('\') -ge 0) {
			$SQLDBName = $SQLDBName.Split("\")[1]
		}

		Write-Log -Message "--------------- getting sql server info --------------------" -LogFile $logfile
		Write-Log -Message "SQLServerName....: $SQLServerName" -LogFile $logfile
		Write-Log -Message "SQLServiceName...: $SQLServiceName" -LogFile $logfile
		Write-Log -Message "SQLPort..........: $SQLPort" -LogFile $logfile
		Write-Log -Message "SQLDBName........: $SQLDBName" -LogFile $logfile

		if (!(Invoke-DbaQuery -SqlInstance $SQLServerName -Database $SQLDBName -Query  "select * from v_r_system" -EnableException -ErrorAction SilentlyContinue)) {
			Write-Log -Message "ERROR / Permission Denied on connection to SQL Server instance" -Category Error -Severity 3 -ShowMsg
			exit
		}
		$arrServers = @()
		$WMIServers = Get-CmWmiObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where NetworkOSPath not like '%.microsoft.com' and Type in (1,2,4,8)" -ComputerName $SmsProvider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
		foreach ($WMIServer in $WMIServers) {
			$arrServers += $WMIServer.NetworkOSPath -replace '\\', ''
		}
		if ($arrServers.Count -gt 0) {
			Write-Log -Message $("Servers discovered: " + $arrServers -join(", ")) -LogFile $LogFile
		}
		else {
			Write-Log -Message "no servers discovered." -LogFile $LogFile
		}
		Write-Log -Message "----------------- creating reporting data table ------------------" -LogFile $LogFile
		$Fields = @("TableName", "XMLFile")
		$ReportTable = New-CmDataTable -TableName $tableName -Fields $Fields

		Write-Log -Message "----------------- creating configuration data table ------------------" -LogFile $LogFile
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
		$outfile = Join-Path -Path $reportFolder -ChildPath "config.xml"
		Write-Log -Message "Exporting XML to $outfile" -LogFile $LogFile
		, $ConfigTable | Export-Clixml -Path $outfile

		$sqlConn = Get-SQLServerConnection -SQLServer "$SQLServerName,$SQLPort" -DBName $SQLDBName
		$sqlConn.Open()

		Write-Log -Message "SQL Query: Creating Functions" -LogFile $LogFile
		$functionsSQLQuery = New-CMHTempSQLfunctions
		$SqlCommand = $sqlConn.CreateCommand()
		$SqlCommand.CommandTimeOut = 0
		$SqlCommand.CommandText = $functionsSQLQuery
		try {
			$SqlCommand.ExecuteNonQuery() | Out-Null
		}
		catch {}
		$SqlCommand = $null
		#$arrSites = @()
		[System.Collections.Generic.List[PSObject]]$arrSites = @()

		$SqlCommand = $sqlConn.CreateCommand()
		$executionquery = "select distinct st.SiteCode, (select top 1 srl2.ServerName from v_SystemResourceList srl2 where srl2.RoleName = 'SMS Provider' and srl2.SiteCode = st.SiteCode) as ServerName from v_Site st"

		Write-Log -Message "sql query........: `n$executionquery" -LogFile $LogFile
		$SqlCommand.CommandTimeOut = 0
		$SqlCommand.CommandText = $executionquery

		Write-Log -Message "--------------- querying database --------------------"	-Severity 1 -LogFile $LogFile
		Write-Log -Message "processing query to sql data adapter..." -LogFile $LogFile
		$DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
		$dataset     = New-Object System.Data.Dataset
		try {
			$DataAdapter.Fill($dataset) | Out-Null
		} catch {
			Write-Log -Message "uh oh! something just blew up, phasers were set to kill... review the log output!" -Severity 3 -LogFile $LogFile
			break
		}
		Write-Log -Message "info.............: data adapter is looking good!" -LogFile $LogFile
		foreach($row in $dataset.Tables[0].Rows) {
			#$arrSites += "$($row.SiteCode)@$($row.ServerName)"
			$arrSites.Add("$($row.SiteCode)@$($row.ServerName)")
		}
		Write-Log -Message $("Sites discovered: " + $arrSites -join(", ")) -LogFile $LogFile
		$SqlCommand = $null

		## section 1 = information that needs be collected against each site
		Write-Log -Message "Phase 1 of 6" -LogFile $logfile -ShowMsg
		foreach ($Site in $arrSites) {
			$arrSiteInfo = $Site.Split("@")
			$PortInformation = Get-CmWmiObject -query "select * from SMS_SCI_Component where FileType=2 and ItemName='SMS_MP_CONTROL_MANAGER|SMS Management Point' and ItemType='Component' and SiteCode='$($arrSiteInfo[0])'" -NameSpace "Root\SMS\Site_$SiteCodeNamespace" -ComputerName $smsprovider -LogFile $logfile
			foreach ($portinfo in $PortInformation) {
				$HTTPport  = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISPortsList"}).Value1
				$HTTPSport = ($portinfo.Props | Where-Object {$_.PropertyName -eq "IISSSLPortsList"}).Value1
			}
			Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '1' -sqlConn $sqlConn -SiteCode $arrSiteInfo[0] -NumberOfDays $NumberOfDays -ServerName $arrSiteInfo[1] -ReportTable $ReportTable -LogFile $logfile
		} # foreach

		## section 2 = information that needs be collected against each computer. should not be site specific. query will run only against the higher site in the hierarchy
		Write-Log -Message "Phase 2 of 6" -LogFile $logfile -ShowMsg
		foreach ($server in $arrServers) {
			Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '2' -sqlConn $sqlConn -siteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $server -ReportTable $ReportTable -LogFile $logfile
		}

		## section 3 = database analisys information, running on all sql servers in the hierarchy. should not be site specific as it connects to the "master" database
		Write-Log -Message "Phase 3 of 6" -LogFile $logfile -ShowMsg
		$DBServers = Get-CmWmiObject -Query "select distinct NetworkOSPath from SMS_SCI_SysResUse where RoleName = 'SMS SQL Server'" -ComputerName $smsprovider -NameSpace "root\sms\site_$SiteCodeNamespace" -LogFile $logfile
		foreach ($DB in $DBServers) {
			$DBServerName = $DB.NetworkOSPath.Replace('\',"")
			Write-Log -Message ("Analysing SQLServer: $DBServerName") -LogFile $LogFile
			if ($SQLServerName.ToLower() -eq $DBServerName.ToLower()) {
				$tmpConnection = $sqlConn
			} else {
				$tmpConnection = Get-SQLServerConnection -SQLServer "$DBServerName,$SQLPort" -DBName "master"
				$tmpConnection.Open()
			}
			try {
				Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '3' -sqlConn $tmpConnection -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ServerName $DBServerName -ReportTable $ReportTable -LogFile $logfile
			} finally {
				if ($SQLServerName.ToLower() -ne $DBServerName.ToLower()) { $tmpConnection.Close()  }
			}
		} # foreach

		## Section 4 = Database analysis against whole SCCM infrastructure, query will run only against top SQL Server
		Write-Log -Message "Phase 4 of 6" -LogFile $logfile -ShowMsg
		Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '4' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile

		## Section 5a = summary information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
		Write-Log -Message "Phase 5 of 6" -LogFile $logfile -ShowMsg
		Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile

		## Section 5b = detailed information against whole SCCM infrastructure. query will run only against the higher site in the hierarchy
		Write-Log -Message "info.............: entering section 5b" -LogFile $LogFile
		Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '5' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -Detailed -LogFile $logfile

		## Section 6 = troubleshooting information
		Write-Log -Message "Phase 6 of 6" -LogFile $logfile -ShowMsg
		Export-ReportSection -HealthCheckXML $HealthCheckXML -Section '6' -sqlConn $sqlConn -SiteCode $SiteCodeNamespace -NumberOfDays $NumberOfDays -ReportTable $ReportTable -LogFile $logfile
	} catch {
		Write-Log -Message "ERROR/EXCEPTION: general unhandled exception" -LogFile $LogFile
		Write-Log -Message "The following error occurred, no further action taken. Dying dying dying, uhhhh.... She's dead, Jim." -LogFile $LogFile
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		Write-Log -Message "Error $errorCode : $errorMessage" -LogFile $LogFile
		$Error.Clear()
	} finally {
		#close sql connection
		Write-Host "Finishing up"
		Write-Log -Message "info.............: closing SQL connection" -LogFile $LogFile
		if ($null -ne $sqlConn) {
			Write-Log -Message "info.............: deleting temp SQL functions" -LogFile $LogFile
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
		if ($null -ne $ReportTable) {
			, $ReportTable | Export-CliXml -Path ($reportFolder + 'report.xml')
		}

		if ($bLogValidation -eq $false) {
			Write-Host "Ending HealthCheck CollectData"
			Write-Log -Message "info.............: log validation not enabled" -LogFile $LogFile
		} else {
			Write-Log -Message "info.............: ending HealthCheck CollectData" -LogFile $LogFile
		}
	}
	$RunTime  = Get-TimeOffset -StartTime $StartTime
	Write-Output "Data collection process completed. Total runtime: $RunTime (hh`:mm`:ss)"
	Write-Log -Message "---------------- FINISHED PROCESSING ------------------" -LogFile $LogFile
}
