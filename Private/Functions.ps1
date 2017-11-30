<#
.SYNOPSIS
	Another stupid custom log writing function
.NOTES
	1.0.3 - 11/29/2017 - David Stein
#>

Function Write-Log {
    param (
		[parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[String] $Message,
        [parameter(Mandatory=$False)][int] $Severity = 1,
        [parameter(Mandatory=$False)][string] $LogFile = '',
        [parameter(Mandatory=$False)][switch] $ShowMsg
        
    )
    #$TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
    #$Date  = Get-Date -Format "HH:mm:ss.fff"
    #$Date2 = Get-Date -Format "MM-dd-yyyy"
    if (($logfile -ne $null) -and ($logfile -ne '')) {
		#"<![LOG[$Message]LOG]!><time=`"$date+$($TimeZoneBias.Bias)`" date=`"$date2`" component=`"$component`" context=`"`" type=`"$severity`" thread=`"`" file=`"`">" | 
		#	Out-File -FilePath $logfile -Append -NoClobber -Encoding Default
		switch ($Severity) {
			1 {$Category='Info'; break}
			2 {$Category='Warning'; break}
			3 {$Category='Error'; break}
		}
		$Msg = "$(Get-Date -f 'yyyy-M-dd HH:mm:ss')  $Category  $Message"
		$Msg | Out-File -FilePath $LogFile -Append -NoClobber -Encoding Default
    }
    if ($showmsg) {
        switch ($Severity) {
            3 { Write-Host $Message -ForegroundColor Red }
            2 { Write-Host $Message -ForegroundColor Yellow }
            1 { Write-Host $Message }
        }
    }
	else {
		Write-Verbose $Msg
	}
}

function Test-Powershell64bit {
    Write-Output ([IntPtr]::size -eq 8)
}

Function Test-Folder {
    param (
        [String] $Path,
        [bool] $Create = $true
    )
    if (Test-Path -Path $Path) {
		return $true
	}
    elseif ($Create -eq $true) {
        try {
            New-Item ($Path) -Type Directory -Force | Out-Null
            Write-Output $true
        }
        catch {
            Write-Output $false
        }
    }
    else {
		Write-Output $false
	}
}

Function Test-RegistryExist {
    param (
		$ComputerName,
		$LogFile = '' ,
		$KeyName,
		$AccessType = 'LocalMachine'
    )
	Write-Log -Message "Testing registry key from $($ComputerName), $($AccessType), $($KeyName)" -LogFile $logfile
    try {
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
		$RegKey = $Reg.OpenSubKey($KeyName)
		$return = ($RegKey -ne $null)
    }
    catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
    }
    Write-Output $return
}

function Test-Admin { 
	$identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent() 
	$principal = New-Object System.Security.Principal.WindowsPrincipal($identity) 
	$admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator 
	$principal.IsInRole($admin) 
} 

function Get-XmlUrlContent {
    param (
        [parameter(Mandatory=$True, HelpMessage="Target URL")]
        [ValidateNotNullOrEmpty()]
        [string] $Url
	)
	Write-Log -Message "reading data from remote file: $Url" -Severity 1 -LogFile $logfile
    $content = ""
    try {
		[xml]$content = ((New-Object System.Net.WebClient).DownloadString($Url))
    }
    catch {}
    if ($content -ne "") {
        $lines = $content -split "`n"
        $result = ""
        for ($i = 1; $i -lt $lines.count; $i++) {
            $result += $lines[$i] + "`n"
        }
    }
    Write-Output $result
}

function Set-ReplaceString {
    param (
	    [string] $Value,
	    [string] $SiteCode,
	    [int] $NumberOfDays = "",
		[string] $ServerName = "",
		[bool] $Space = $true
	)
	$return = $value
    $date = Get-Date
	if ($space) {	
		$return = $return -replace "\r\n", " " 
		$return = $return -replace "\r", " " 
		$return = $return -replace "\n", " " 
		$return = $return -replace "\s", " " 
		$return = $return -replace "\s{2}\b"," "
	}
	$return = $return -replace "@@SITECODE@@",$SiteCode
	$return = $return -replace "@@STARTMONTH@@",$date.ToString("01/MM/yyyy")
	$return = $return -replace "@@TODAYMORNING@@",$date.ToString("yyyy/MM/dd")
	$return = $return -replace "@@NUMBEROFDAYS@@",$NumberOfDays
	$return = $return -replace "@@SERVERNAME@@",$ServerName
	if ($space) {
		while (($return.IndexOf("  ") -ge 0)) { $return = $return -replace "  ", " " }
	}
	Write-Output $return
}

Function Get-RegistryValue {
    param (
        [String] $ComputerName,
        [string] $LogFile = '' ,
        [string] $KeyName,
        [string] $KeyValue,
        [string] $AccessType = 'LocalMachine'
    )
    Write-Log -Message "Getting registry value from $($ComputerName), $($AccessType), $($keyname), $($keyvalue)" -Severity 1 -LogFile $logfile
    try {
        $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
        $RegKey= $Reg.OpenSubKey($keyname)
	    if ($RegKey -ne $null) {
		    try { $return = $RegKey.GetValue($keyvalue) }
		    catch { $return = $null }
	    }
	    else { $return = $null }
        
        Write-Log -Message "Value returned $return" -Severity 1 -LogFile $logfile
    }
    catch {
        $return = "ERROR: Unknown"
        $Error.Clear()
    }
    Write-Output $return
}

Function ReportSection {
    param (
	    $HealthCheckXML,
		$Section,
		$SqlConn,
		[string] $SiteCode,
		$NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		$ReportTable,
		[switch] $Detailed
	)
	Write-Log -Message "[function: ReportSection]" -LogFile $logfile
	if ($Detailed) { 
        Write-Log -Message "[detailed = True]" -LogFile $logfile
		Write-Log -Message "-----------------------------------------------------------"  -LogFile $logfile
		Write-Log -Message "**** Starting Section $Section with [Detailed] = $($detailed.ToString())"  -LogFile $logfile
		Write-Log -Message "-----------------------------------------------------------"  -LogFile $logfile
	}
	
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
        if ($healthCheck.IsTextOnly.ToLower() -eq 'true') { continue }
        if ($healthCheck.IsActive.ToLower() -ne 'true') { continue }
		if ($healthCheck.Section.ToLower() -ne $Section) { continue }	
		$sqlquery  = $healthCheck.SqlQuery
        $tablename = (Set-ReplaceString -Value $healthCheck.XMLFile -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $servername)
        $xmlTableName = $healthCheck.XMLFile
        if ($Section -eq 5) {
            if (!($Detailed)) { 
                $tablename += "summary" 
                $xmlTableName += "summary"
                $gbfiels = ""
                foreach ($field in $healthCheck.Fields.Field) {
                    if ($field.groupby -in ("2")) {
                        if ($gbfiels.Length -gt 0) { $gbfiels += "," }
                        $gbfiels += $field.FieldName
                    }
                }
                $sqlquery = "select $($gbfiels), count(1) as Total from ($($sqlquery)) tbl group by $($gbfiels)"
            } 
            else { 
                $tablename += "detail"
                $xmlTableName += "detail"
                $sqlquery = $sqlquery -replace "select distinct", "select"
                $sqlquery = $sqlquery -replace "select", "select distinct"
            }
        }
    	$filename = $reportFolder + $tablename + '.xml'
		$row = $ReportTable.NewRow()
    	$row.TableName = $xmlTableName
    	$row.XMLFile = $tablename + ".xml"
    	$ReportTable.Rows.Add($row)
		Write-Log -Message ("XMLfile... $filename")  -LogFile $logfile
		Write-Log -Message ("Section... $Section")  -LogFile $logfile
		Write-Log -Message ("Table..... $TableName - Information...Starting") -LogFile $logfile
		Write-Log -Message ("Type...... $($healthCheck.querytype)") -LogFile $logfile
		try {
			switch ($healthCheck.querytype.ToLower()) {
				'mpconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mplist' | Out-Null}
				'mpcertconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mpcert' | Out-Null}
				'sql' { Get-SQLData -sqlConn $sqlConn -SQLQuery $sqlquery -FileName $fileName -TableName $tablename -siteCode $siteCode -NumberOfDays $NumberOfDays -servername $servername -healthcheck $healthCheck -logfile $logfile -section $section -detailed $detailed | Out-Null}
				'baseosinfo' { Write-BaseOSInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'diskinfo' { Write-DiskInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'networkinfo' { Write-NetworkInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'rolesinstalled' { Write-RolesInstalled -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile | Out-Null}
				'servicestatus' { Write-ServiceStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'hotfixstatus' { 
                    if (-not $NoHotfix) {
                        Write-HotfixStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null
                    }
                }
           		default {}
			}
		}
		catch {
			$errorMessage = $Error[0].Exception.Message
			$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred..." -Severity 3 -LogFile $logfile
			Write-Log -Message "Error $errorCode : $errorMessage connecting to $servername" -Severity 3 -LogFile $logfile
			$Error.Clear()
		}
		Write-Log -Message ("$tablename Information...Done") -LogFile $logfile
    }
	Write-Log -Message "End Section $section"
}

Function Set-FormatedValue {
    param (
	    $Value,
	    [string] $Format,
		[string] $SiteCode
	)
	Write-Log -Message "[function: Set-FormatedValue]" -LogFile $logfile
	Write-Log -Message "  [format = $Format]" -LogFile $logfile
	Write-Log -Message "  [sitecode = $SiteCode]" -LogFile $logfile
	if ($Value -eq $null) {
		Write-Log -Message "  [value = NULL]" -LogFile $logfile
	}
	else {
		Write-Log -Message "  [value = $Value]" -LogFile $logfile
	}
	switch ($format.ToLower()) {
		'schedule' {
			$schedule_Class = [wmiclass]""
			$schedule_class.psbase.path = "\\$($smsprovider)\root\sms\site_$($SiteCodeNamespace):SMS_ScheduleMethods"
			$schedule = ($schedule_class.ReadFromString($value)).TokenData
			if ($schedule.DaySpan -ne 0) { $return = ($schedule.DaySpan * 24 * 60) }
			elseif ($schedule.HourSpan -ne 0) { $return = ($schedule.HourSpan * 60) }
			elseif ($schedule.MinuteSpan -ne 0) { $return = ($schedule.MinuteSpan) }
			return $return
			break
		}
        'alertsname' {
			if ($value -eq $null) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'$databasefreespacewarning' {
						$return = 'Low free space alert for database on site'
						break
					}
					'$sumcompliance2updategroupdeploymentname' {
						$return = 'Low deployment success rate alert of update group'
						break
					}
					default {
						$return = $value
						break
					}
				}
			}
            return $return
            break
        }
        'alertsseverity' {
			if ($value -eq $null) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'1' {
						$return = 'Error'
						break
					}
					'2' {
						$return = 'Warning'
						break
					}
					'3' {
						$return = 'Informational'
						break
					}
					default {
						$return = 'Unknown'
						break
					}
				}
			}
            return $return
            break
        }
        'alertstypeid' {
            switch ($value.ToString().ToLower()) {
                '12' {
                    $return = 'Update group deployment success'
                    break
                }
                '25' {
                    $return = 'Database free space warning'
                    break
                }
                '31' {
                    $return = 'Malware detection'
                    break
                }
                default {
                    $return = $value
                    break
                }
            }
            Write-Output $return
            break
        }
		'messagesolution' {
			Write-Log -Message "[messagesolution] convert to string" -LogFile $logfile
			if ($value -ne $null) {
				$return = $value.ToString()
			}
			Write-Output $return
			break
		}
		default {
			Write-Output $value
			break
		}
	}
}

Function Get-SQLData {
    param (
	    [parameter(Mandatory=$True)]
			$sqlConn,
	    [parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[string] $SQLQuery,
	    [parameter(Mandatory=$False)]
			[string] $FileName,
	    [parameter(Mandatory=$False)]
			[string] $TableName,
	    [parameter(Mandatory=$False)]
			[string] $SiteCode,
	    [parameter(Mandatory=$False)]
			$NumberOfDays,
	    [parameter(Mandatory=$False)]
			$LogFile,
		[parameter(Mandatory=$False)]
			[string] $ServerName,
		[parameter(Mandatory=$False)]
			[bool] $ContinueOnError = $true,
		[parameter(Mandatory=$False)]
			$HealthCheck,
        [parameter(Mandatory=$False)]
			$Section,
        [parameter(Mandatory=$False)]
			[switch] $Detailed
	)
	Write-Log -Message "[function: Get-SQLData]"
	if ($Detailed) { 
        Write-Log -Message "  [detailed = True]" -LogFile $logfile
    }
    try {
        $SqlCommand = $sqlConn.CreateCommand()
		$logQuery       = Set-ReplaceString -value $SQLQuery -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $ServerName
		$executionquery = Set-ReplaceString -value $SQLQuery -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $ServerName -space $false
        Write-Log -Message "SQL Query...`n$executionquery" -LogFile $logfile
	    Write-Log -Message "Log Query...`n$logQuery" -LogFile $logfile
        $SqlCommand.CommandTimeOut = 0
        $SqlCommand.CommandText = $executionquery
        $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
        $dataset     = New-Object System.Data.Dataset
        $DataAdapter.Fill($dataset)
		if (($dataset.Tables.Count -eq 0) -or ($dataset.Tables[0].Rows.Count -eq 0)) { 
			Write-Log -Message "SQL Query returned 0 records" -LogFile $logfile
			Write-Log -Message "Table $tablename is empty. No file output to $filename ..." -LogFile $logfile
		}
		else {
			Write-Log -Message "SQL Query returned $($dataset.Tables[0].Rows.Count) records"
			foreach ($field in $healthCheck.Fields.Field) {
				Write-Log -Message ("   field = $($Field.FieldName) description = $($Field.Description)") -LogFile $logfile
                if ($section -eq 5) {
                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                    elseif (($detailed -eq $false) -and ($field.groupby -notin ('2','3'))) { continue }
                }
				if ($field.format -ne "") {
					Write-Log -Message "   custom format specified for this attribute: $($Field.Format)" -LogFile $logfile
					foreach ($row in $dataset.Tables[0].Rows) {
						#$fn = $field.FieldName
						$tempx = Set-FormatedValue -Value $row.$($field.FieldName) -Format $field.format -SiteCode $SiteCode
						try {
							$row.$($field.FieldName) = $tempx
						}
						catch {
							$row
							break
						}
					}
				}
			}
			Write-Log -Message "Export: Exporting xml data to $filename" -LogFile $logfile
        	, $dataset.Tables[0] | Export-CliXml -Path $filename
		}
    }
    catch {
        $errorMessage = $Error[0].Exception.Message
        $errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
        if ($continueonerror -eq $false) { 
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred (stop)." -Severity 3 -LogFile $logfile
		}
        else { 
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred (continue)." -Severity 3 -LogFile $logfile
		}
        Write-Log -Message "Error $errorCode : $errorMessage connecting to $ServerName" -Severity 3 -LogFile $logfile
	    $Error.Clear()
		Write-Log -Message "Unable to update file: $filename" -Severity 2 -LogFile $logfile
        if ($continueonerror -eq $false) {
			Throw "Error $errorCode : $errorMessage connecting to $ServerName"
		}
	}
}

function New-CmDataTable {
    param (
		[string] $TableName,
	    [String[]] $Fields
    )
    Write-Log -Message "[function: New-CmDataTable]" -LogFile $logfile
	$DataTable = New-Object System.Data.DataTable "$tableName"
	foreach ($field in $fields) {
		$col = New-Object System.Data.DataColumn "$field",([string])
		$DataTable.Columns.Add($col)
	}
	,$DataTable
}

Function Write-BaseOSInfo {
    param (
	    [string] $FileName,
	    [string] $TableName,
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-baseosinfo]" -LogFile $logfile
    $WMIOS = Get-CmWmiObject -Class "win32_operatingsystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($WMIOS -eq $null) { return }	
    $WMICS = Get-CmWmiObject -Class "win32_computersystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
	$WMIProcessor = Get-CmWmiObject -Class "Win32_processor" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    $WMITimeZone  = Get-CmWmiObject -Class "Win32_TimeZone" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    ##AV Information
    $avInformation = $null
    $AVArray = @("McAfee Security@McShield", "Symantec Endpoint Protection@symantec antivirus", "Sophos Antivirus@savservice", "Avast!@aveservice", "Avast!@avast! antivirus", "Immunet Protect@immunetprotect", "F-Secure@fsma", "AntiVir@antivirservice", "Avira@avguard", "F-Protect@fpavserver", "Panda Security@pshost", "Panda AntiVirus@pavsrv", "BitDefender@bdss", "ArcaBit/ArcaVir@abmainsv", "IKARUS@ikarus-guardx", "ESET Smart Security@ekrn", "G Data Antivirus@avkproxy", "Kaspersky Lab Antivirus@klblmain", "Symantec VirusBlast@vbservprof", "ClamAV@clamav", "Vipre / GFI managed AV@SBAMSvc", "Norton@navapsvc", "Kaspersky@AVP", "Windows Defender@windefend", "Windows Defender/@MsMpSvc", "Microsoft Security Essentials@msmpeng")

    foreach ($av in $AVArray) {
        $info = $av.Split("@")
        if ((Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $info[1]).ToString().Tolower().Indexof("error") -lt 0) {
            $avInformation = $info[0]
            break
        }
    }
    $OSProcessorArch = $WMIOS.OSArchitecture
    if ($OSProcessorArch -ne $null) {
	    switch ($OSProcessorArch.ToUpper() ) {
		    "AMD64" {$ProcessorArchDisplay = "64-bit"}
			"i386"  {$ProcessorArchDisplay = "32-bit"}
			"IA64"  {$ProcessorArchDisplay = "64-bit - Itanium"}
			default {$ProcessorArchDisplay = $OSProcessorArch }
	    }
	} 
    else { 
        $ProcessorArchDisplay = "" 
    }
    $LastBootUpTime = $WMIOS.ConvertToDateTime($WMIOS.LastBootUpTime)
    $LocalDateTime  = $WMIOS.ConvertToDateTime($WMIOS.LocalDateTime)
    $numProcs = 0
	$ProcessorType = ""
	$ProcessorName = ""
	$ProcessorDisplayName= ""
	foreach ($WMIProc in $WMIProcessor) {
		$ProcessorType = $WMIProc.Manufacturer
		switch ($WMIProc.NumberOfCores) {
			1 {$numberOfCores = "single core"}
			2 {$numberOfCores = "dual core"}
			4 {$numberOfCores = "quad core"}
			$null {$numberOfCores = "single core"}
			default { $numberOfCores = $WMIProc.NumberOfCores.ToString() + " core" } 
		}
		switch ($WMIProc.Architecture) {
			0 {$CpuArchitecture = "x86"}
			1 {$CpuArchitecture = "MIPS"}
			2 {$CpuArchitecture = "Alpha"}
			3 {$CpuArchitecture = "PowerPC"}
			6 {$CpuArchitecture = "Itanium"}
			9 {$CpuArchitecture = "x64"}
		}
		if ($ProcessorDisplayName.Length -eq 0) { 
			$ProcessorDisplayName = " " + $numberOfCores + " $CpuArchitecture processor " + $WMIProc.Name
		}
        else {
			if ($ProcessorName -ne $WMIProc.Name) { 
				$ProcessorDisplayName += "/ " + " " + $numberOfCores + " $CpuArchitecture processor " + $WMIProc.Name
			}
		}
		$numProcs += 1
		$ProcessorName = $WMIProc.name
	}
	$ProcessorDisplayName = "$numProcs" + $ProcessorDisplayName
    if ($WMICS.DomainRole -ne $null) {
		switch ($WMICS.DomainRole) {
			0 {$RoleDisplay = "Workstation"}
			1 {$RoleDisplay = "Member Workstation"}
			2 {$RoleDisplay = "Standalone Server"}
			3 {$RoleDisplay = "Member Server"}
			4 {$RoleDisplay = "Backup Domain Controller"}
			5 {$RoleDisplay = "Primary Domain controller"}
            default: {$RoleDisplay = "unknown, $($WMICS.DomainRole)"}
		}
	}
	$Fields = @("ComputerName","OperatingSystem","ServicePack","Version","Architecture","LastBootTime","CurrentTime","TotalPhysicalMemory","FreePhysicalMemory","TimeZone","DaylightInEffect","Domain","Role","Model","NumberOfProcessors","NumberOfLogicalProcessors","Processors","AntiMalware")
	$BaseOSInfoTable = New-CmDataTable -TableName $tableName -Fields $Fields
	$row = $BaseOSInfoTable.NewRow()
	$row.ComputerName = $ServerName
	$row.OperatingSystem = $WMIOS.Caption
	$row.ServicePack = $WMIOS.CSDVersion
	$row.Version = $WMIOS.Version
	$row.Architecture = $ProcessorArchDisplay
	$row.LastBootTime = $LastBootUpTime.ToString()
	$row.CurrentTime  = $LocalDateTime.ToString()
	$row.TotalPhysicalMemory = ([string]([math]::Round($($WMIOS.TotalVisibleMemorySize/1MB), 2)) + " GB")
	$row.FreePhysicalMemory = ([string]([math]::Round($($WMIOS.FreePhysicalMemory/1MB), 2)) + " GB")
	$row.TimeZone = $WMITimeZone.Description
	$row.DaylightInEffect = $WMICS.DaylightInEffect
	$row.Domain = $WMICS.Domain
	$row.Role   = $RoleDisplay
	$row.Model  = $WMICS.Model
	$row.NumberOfProcessors = $WMICS.NumberOfProcessors
	$row.NumberOfLogicalProcessors = $WMICS.NumberOfLogicalProcessors
	$row.Processors = $ProcessorDisplayName
    if ($avInformation -ne $null) { $row.AntiMalware = $avInformation }
    else { $row.AntiMalware = "Antimalware software not detected" }
    $BaseOSInfoTable.Rows.Add($row)
    , $BaseOSInfoTable | Export-CliXml -Path ($filename)
}

Function Write-DiskInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-diskinfo]" -LogFile $logfile
    $DiskList = Get-CmWmiObject -Class "Win32_LogicalDisk" -Filter "DriveType = 3" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($DiskList -eq $null) { return }
	$Fields=@("DeviceID","Size","FreeSpace","FileSystem")
	$DiskDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($Disk in $DiskList) {
		$row = $DiskDetails.NewRow()
		$row.DeviceID = $Disk.DeviceID
		$row.Size = ([int](($Disk.Size) / 1024 / 1024 / 1024)).ToString()
		$row.FreeSpace = ([int](($Disk.FreeSpace) / 1024 / 1024 / 1024)).ToString()
		$row.FileSystem = $Disk.FileSystem
	    $DiskDetails.Rows.Add($row)
    }
    , $DiskDetails | Export-CliXml -Path ($filename)
}

function Write-NetworkInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-networkinfo]" -LogFile $logfile
    $IPDetails = Get-CmWmiObject -Class "Win32_NetworkAdapterConfiguration" -Filter "IPEnabled = true" -ComputerName $servername -logfile $logfile -continueonerror $continueonerror
    if ($IPDetails -eq $null) { return }
	$Fields = @("IPAddress","DefaultIPGateway","IPSubnet","MACAddress","DHCPEnabled")
	$NetworkInfoTable = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($IPAddress in $IPDetails) {
		$row = $NetworkInfoTable.NewRow()
		$row.IPAddress = ($IPAddress.IPAddress -join ", ")
		$row.DefaultIPGateway = ($IPAddress.DefaultIPGateway -join ", ")
		$row.IPSubnet = ($IPAddress.IPSubnet -join ", ")
		$row.MACAddress = $IPAddress.MACAddress
		if ($IPAddress.DHCPEnable -eq $true) { $row.DHCPEnabled = "TRUE" } else { $row.DHCPEnabled = "FALSE" }
	    $NetworkInfoTable.Rows.Add($row)
    }
    , $NetworkInfoTable | Export-CliXml -Path ($filename)
}

function Write-RolesInstalled {
    param (
	    [string] $FileName,
	    [string] $TableName,
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogfFle,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-rolesinstalled]" -LogFile $logfile
    $WMISMSListRoles = Get-CmWmiObject -Query "select distinct RoleName from SMS_SCI_SysResUse where NetworkOSPath = '\\\\$Servername'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
    $SMSListRoles = @()
    foreach ($WMIServer in $WMISMSListRoles) { $SMSListRoles += $WMIServer.RoleName }
    $DPProperties = Get-CmWmiObject -Query "select * from SMS_SCI_SysResUse where RoleName = 'SMS Distribution Point' and NetworkOSPath = '\\\\$Servername' and SiteCode = '$SiteCode'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
 	$Fields = @("SiteServer", "IIS", "SQLServer", "DP", "PXE", "MultiCast", "PreStaged", "MP", "FSP", "SSRS", "EP", "SUP", "AI", "AWS", "PWS", "SMP", "Console", "Client", "CPC", "DWP", "DMP")
	$RolesInstalledTable = New-CmDataTable -TableName $tableName -Fields $Fields
	$row = $RolesInstalledTable.NewRow()
	$row.SiteServer = ($SMSListRoles -contains 'SMS Site Server').ToString()
	$row.SQLServer  = ($SMSListRoles -contains 'SMS SQL Server').ToString()
	$row.DP = ($SMSListRoles -contains 'SMS Distribution Point').ToString()
	if ($DPProperties -eq $null) {
		$row.PXE = "False"
		$row.MultiCast = "False"
		$row.PreStaged = "False"
	}
	else {
		$row.PXE = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "IsPXE"}).Value -eq 1).ToString()
		$row.MultiCast = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "IsMulticast"}).Value -eq 1).ToString()
		$row.PreStaged = (($DPProperties.Props | Where-Object {$_.PropertyName -eq "PreStagingAllowed"}).Value -eq 1).ToString()
	}
	$row.MP   = ($SMSListRoles -contains 'SMS Management Point').ToString()
	$row.FSP  = ($SMSListRoles -contains 'SMS Fallback Status Point').ToString()
	$row.SSRS = ($SMSListRoles -contains 'SMS SRS Reporting Point').ToString()
	$row.EP   = ($SMSListRoles -contains 'SMS Endpoint Protection Point').ToString()
	$row.SUP  = ($SMSListRoles -contains 'SMS Software Update Point').ToString()
	$row.AI   = ($SMSListRoles -contains 'AI Update Service Point').ToString()
	$row.AWS  = ($SMSListRoles -contains 'SMS Application Web Service').ToString()
	$row.PWS  = ($SMSListRoles -contains 'SMS Portal Web Site').ToString()
	$row.SMP  = ($SMSListRoles -contains 'SMS State Migration Point').ToString()
	# added in 0.64
	$row.CPC  = ($SMSListRoles -contains 'SMS Cloud Proxy Connector').ToString()
	$row.DWP  = ($SMSListRoles -contains 'Data Warehouse Service Point').ToString()
	$row.DMP  = ($SMSListRoles -contains 'SMS Dmp Connector').ToString()
	# other roles as of build 1702
	<#
	SMS Device Management Point
	SMS System Health Validator
	SMS Multicast Service Point
	SMS AMT Service Point
	SMS Enrollment Server
	SMS Enrollment Web Site
	SMS Notification Server
	SMS Certificate Registration Point
	SMS DM Enrollment Service
	#>
	$row.Console = (Test-RegistryExist -ComputerName $servername -Logfile $logfile -KeyName 'SOFTWARE\\Wow6432Node\\Microsoft\\ConfigMgr10\\AdminUI').ToString()
	$row.Client  = (Test-RegistryExist -ComputerName $servername -Logfile $logfile -KeyName 'SOFTWARE\\Microsoft\\CCM\\CCMExec').ToString()
	$row.IIS     = ((Get-RegistryValue -ComputerName $server -Logfile $logfile -KeyName 'SOFTWARE\\Microsoft\\InetStp' -KeyValue 'InstallPath') -ne $null).ToString()
    $RolesInstalledTable.Rows.Add($row)
    , $RolesInstalledTable | Export-Clixml -Path ($filename)
}

Function Get-ServiceStatus {
	param (
		$LogFile,
		[string] $ServerName,
		[string] $ServiceName
    )
    Write-Log -Message "[function: get-servicestatus]" -LogFile $logfile
	Write-Log -Message "  servername = $servername / servicename = $servicename" -LogFile $logfile
    try {
		$service = Get-Service -ComputerName $servername | Where-Object {$_.Name -eq $servicename}
		if ($service -ne $null) { $return = $service.Status }
		else  { $return = "ERROR: Not Found" }
		Write-Log -Message "Service status $return" -LogFile $logfile
    }
    catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
    }
    Write-Output $return
}

function Write-MPConnectivity {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $Type = 'mplist'
    )
    Write-Log -Message "[function: write-mpconnectivity]" -LogFile $logfile
 	$Fields = @("ServerName", "HTTPReturn")
	$MPConnectivityTable = New-CmDataTable -TableName $tableName -Fields $Fields

	$MPList = Get-CmWmiObject -query "select * from SMS_SCI_SysResUse where SiteCode = '$SiteCode' and RoleName = 'SMS Management Point'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
	foreach ($MPInformation in $MPList) {
	    $SSLState = ($MPInformation.Props | Where-Object {$_.PropertyName -eq "SslState"}).Value
		$mpname = $MPInformation.NetworkOSPath -replace '\\', ''
	    if ($SSLState -eq 0) {
			$protocol = 'http'
			$port = $HTTPport 
		} 
		else {
			$protocol = 'https'
			$port = $HTTPSport 
		}
	            
		$web = New-Object -ComObject msxml2.xmlhttp
		$url = $protocol + '://' + $mpname + ':' + $port + '/sms_mp/.sms_aut?' + $type
        if ($healthcheckdebug) { Write-Verbose ("URL Connection: $url") }
		$row = $MPConnectivityTable.NewRow()
		$row.ServerName = $mpname
	    try {   
			$web.open('GET', $url, $false)
			$web.send()
			$row.HTTPReturn = $web.status
	    }
	    catch {
			$row.HTTPReturn = "313 - Unable to connect to host"
			$Error.Clear()
	    }
		Write-Log -Message "  Status: $($web.status)" -LogFile $logfile
		$MPConnectivityTable.Rows.Add($row)
	}
    , $MPConnectivityTable | Export-CliXml -Path ($filename)
}

Function Write-HotfixStatus {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $ServerName,
		$ContinueOnError = $true
    )
    Write-Log -Message "[function: write-hotfixstatus]" -LogFile $logfile
    try {         
		$Session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $ServerName))
		$Searcher = $Session.CreateUpdateSearcher()
		$historyCount = $Searcher.GetTotalHistoryCount()
		$return = $Searcher.QueryHistory(0, $historyCount) 
		Write-Log -Message "  Hotfix count: $HistoryCount" -LogFile $logfile
    }
    catch {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		Write-Log -Message "  The following error happen" -LogFile $logfile
		Write-Log -Message "  Error $errorCode : $errorMessage connecting to $ServerName" -LogFile $logfile
		$Error.Clear()
		return
    }
    $Fields = @("Title", "Date")
	$HotfixTable = New-CmDataTable -tablename $tableName -fields $Fields
    foreach ($hotfix in $return) {
		$row = $HotfixTable.NewRow()
		$row.Title = $hotfix.Title
		$row.Date  = $hotfix.Date
		$HotfixTable.Rows.Add($row)
    }
    , $HotfixTable | Export-CliXml -Path ($filename)
}

function Write-ServiceStatus {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		$LogFile,
		[string] $ServerName,
		$ContinueOnError = $true
    )
    Write-Log -Message "[function: write-servicestatus]" -LogFile $logfile

	$SiteInformation = Get-CmWmiObject -query "select Type from SMS_Site where ServerName = '$Server'" -namespace "Root\SMS\Site_$SiteCodeNamespace" -computerName $smsprovider -logfile $logfile
    if ($SiteInformation -ne $null) { $SiteType = $SiteInformation.Type }

    $WMISMSListRoles = Get-CmWmiObject -query "select distinct RoleName from SMS_SCI_SysResUse where NetworkOSPath = '\\\\$Server'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
    $SMSListRoles = @()
    foreach ($WMIServer in $WMISMSListRoles) { $SMSListRoles += $WMIServer.RoleName }
	Write-Log -Message "Roles discovered: " + $SMSListRoles -join(", ") -LogFile $logfile
 	$Fields = @("ServiceName", "Status")
	$ServicesTable = New-CmDataTable -TableName $tableName -Fields $Fields

    if ($SMSListRoles -contains 'AI Update Service Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "AI_UPDATE_SERVICE_POINT"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    if (($SMSListRoles -contains 'SMS Application Web Service') -or ($SMSListRoles -contains 'SMS Distribution Point') -or ($SMSListRoles -contains 'SMS Fallback Status Point') -or ($SMSListRoles -contains 'SMS Management Point') -or ($SMSListRoles -contains 'SMS Portal Web Site')  ) {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "IISADMIN"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "W3SVC"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS Component Server') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_Executive"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS Site Server') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_NOTIFICATION_SERVER"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SMS_SITE_COMPONENT_MANAGER"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
		if ($SiteType -ne 1) {
			$row = $ServicesTable.NewRow()
			$row.ServiceName = "SMS_SITE_VSS_WRITER"
			$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
		    $ServicesTable.Rows.Add($row)
		}
    }
    if ($SMSListRoles -contains 'SMS Software Update Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "WsusService"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS SQL Server') {
		$row = $ServicesTable.NewRow()
		if ($SiteType -ne 1) {		
			$row.ServiceName = "$SQLServiceName"
		}
		else {
			$row.ServiceName = 'MSSQL$CONFIGMGRSEC'
		}
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
		$ServicesTable.Rows.Add($row)
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "SQLWriter"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS SRS Reporting Point') {
		$row = $ServicesTable.NewRow()
		$row.ServiceName = "ReportServer"
		$row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
	    $ServicesTable.Rows.Add($row)
    }
    , $ServicesTable | Export-CliXml -Path ($filename)
}

function Get-CmCredentials {
    try {
        $cred = Get-Credentials
        Write-Log -Message "  Trying username: $($cred.Username)" -LogFile $logfile
        Write-Output $cred
    }
    catch {
        Write-Output $null
    }
}

function Get-CmWmiObject {
    param (
		$Class,
		$Filter = '',
		$Query = '',
		$ComputerName,
		$Namespace = "root\cimv2",
		$LogFile,
		[bool] $ContinueOnError = $false
    )
    if ($query -ne '') { $class = $query }
	Write-Log -Message "  WMI Query: \\$ComputerName\$Namespace, $class, filter: $filter" -LogFile $logfile
    if ($query -ne '') { 
		$WMIObject = Get-WmiObject -Query $query -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
    elseif ($filter -ne '') { 
		$WMIObject = Get-WmiObject -Class $class -Filter $filter -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
    else { 
		$WMIObject = Get-WmiObject -Class $class -NameSpace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}

	if ($WMIObject -eq $null) {
		Write-Log -Message "  WMI Query returned 0) records" -LogFile $logfile
	}
	else {
		$i = 1
		foreach ($obj in $wmiobj) { i++ }
		Write-Log -Message "  WMI Query returned $($i) records" -LogFile $logfile
	}
	
    if ($Error.Count -ne 0) {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		if ($ContinueOnError -eq $false) {
            Write-Log -Message "  The following error occurred, no futher action taken" -Severity 3 -Logfile $logfile
        }
		else { 
            Write-Error "The following error occurred"
        }
		Write-Log -Message "  Error $errorCode : $errorMessage connecting to $ComputerName" -LogFile $logfile
		$Error.Clear()
		if ($ContinueOnError -eq $false) { 
            Throw "Error $errorCode : $errorMessage connecting to $ComputerName" 
        }
    }
    Write-Output $WMIObject
}

Function Get-SQLServerConnection {
    param (
		[parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[string] $SQLServer,
		[parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[string] $DBName
    )
    Try {
		$conn = New-Object System.Data.SqlClient.SqlConnection
		$conn.ConnectionString = "Data Source=$SQLServer;Initial Catalog=$DBName;Integrated Security=SSPI;"
		return $conn
    }
    Catch {
		$errorMessage = $_.Exception.Message
		$errorCode = "0x{0:X}" -f $_.Exception.ErrorCode
		Write-Log -Message "The following error happen, no futher action taken" -LogFile $logfile
		Write-Log -Message "Error $errorCode : $errorMessage connecting to $SQLServer" -Severity 3 -LogFile $logfile
		$Error.Clear()
		Throw "Error $errorCode : $errorMessage connecting to $SQLServer"
    }
}

#----------------------------------
function Get-MessageInformation {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.Message | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null) {
        Write-Output "Unknown Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}

function Get-MessageSolution {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.MessageSolution | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null)	{ 
        Write-Output "There is no known possible solution for Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}

function Write-WordText {
    param (
		$WordSelection,
		[string] $Text    = "",
		[string] $Style   = "No Spacing",
		$Bold    = $false,
		$NewLine = $false,
		$NewPage = $false
	)
	$texttowrite = ""
	$wordselection.Style = $Style
    if ($bold) { $wordselection.Font.Bold = 1 } else { $wordselection.Font.Bold = 0 }
	$texttowrite += $text
	$wordselection.TypeText($text)
	If ($newline) { $wordselection.TypeParagraph() }
	If ($newpage) { $wordselection.InsertNewPage() }
}

Function Set-WordDocumentProperty {
    param (
		$Document,
		$Name,
		$Value
	)
    Write-Log -Message "info: document property [$Name] set to [$Value]" -LogFile $logfile
    $document.BuiltInDocumentProperties($Name) = $Value
}

Function Write-WordReportSection {
    param (
		$HealthCheckXML,
        $Section,
		$Detailed = $false,
        $Doc,
		$Selection,
        $LogFile
	)
	Write-Log -Message "Starting Section $section with detailed as $($detailed.ToString())" -LogFile $logfile
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.Section.tolower() -ne $Section) { continue }
		$Description = $healthCheck.Description -replace("@@NumberOfDays@@", $NumberOfDays)
		if ($healthCheck.IsActive.tolower() -ne 'true') { continue }
        if ($healthCheck.IsTextOnly.tolower() -eq 'true') {
            if ($Section -eq 5) {
                if ($detailed -eq $false) { 
                    $Description += " - Overview" 
                } 
                else { 
                    $Description += " - Detailed"
                }            
            }
			Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
			Continue;
		}
		Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
        $bFound = $false
        $tableName = $healthCheck.XMLFile
        if ($Section -eq 5) {
            if (!($detailed)) { 
                $tablename += "summary" 
            } 
            else { 
                $tablename += "detail"
            }            
        }
		foreach ($rp in $ReportTable) {
			if ($rp.TableName -eq $tableName) {
                $bFound = $true
				Write-Log -Message (" - Exporting $($rp.XMLFile) ...") -LogFile $logfile
				$filename = $rp.XMLFile				
				if ($filename.IndexOf("_") -gt 0) {
					$xmltitle = $filename.Substring(0,$filename.IndexOf("_"))
					$xmltile = ($rp.TableName.Substring(0,$rp.TableName.IndexOf("_")).Replace("@","")).Tolower()
					switch ($xmltile) {
						"sitecode"   { $xmltile = "Site Code: "; break; }
						"servername" { $xmltile = "Server Name: "; break; }
					}
					switch ($healthCheck.WordStyle) {
						"Heading 1" { $newstyle = "Heading 2"; break; }
						"Heading 2" { $newstyle = "Heading 3"; break; }
						"Heading 3" { $newstyle = "Heading 4"; break; }
						default { $newstyle = $healthCheck.WordStyle; break }
					}
					$xmltile += $filename.Substring(0,$filename.IndexOf("_"))
					Write-WordText -WordSelection $selection -Text $xmltile -Style $newstyle -NewLine $true
				}				
				
	            if (!(Test-Path ($reportFolder + $rp.XMLFile))) {
					Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
					Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
					$selection.TypeParagraph()
				}
				else {
					Write-Log -Message "importing XML file: $filename"
					$datatable = Import-CliXml -Path ($reportFolder + $filename)
					$count = 0
					$datatable | Where-Object { $count++ }
					
		            if ($count -eq 0) {
						Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
						Write-Log -Message ("Table: 0 rows") -LogFile $logfile -Severity 2
						$selection.TypeParagraph()
						continue
		            }

					switch ($healthCheck.PrintType.ToLower()) {
						"table" {
                            Write-Log -Message "writing table type: table" -LogFile $logfile
							$Table = $Null
					        $TableRange = $Null
					        $TableRange = $doc.Application.Selection.Range
                            $Columns = 0
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }
							$Table = $doc.Tables.Add($TableRange, $count+1, $Columns)
							$table.Style = $TableStyle
							$i = 1;
							Write-Log -Message ("Table: $count rows and $Columns columns") -LogFile $logfile
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }

								$Table.Cell(1, $i).Range.Font.Bold = $True
								$Table.Cell(1, $i).Range.Text = $field.Description
								$i++
	                        }
							$xRow = 2
							$records = 1
							$y=0
							foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								$i = 1;
								foreach ($field in $HealthCheck.Fields.Field) {
                                    if ($section -eq 5) {
                                        if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                        elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                    }
									$Table.Cell($xRow, $i).Range.Font.Bold = $false
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($row.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($row.$($field.FieldName))
											break ;
										}										
										default {
											$TextToWord = $row.$($field.FieldName);
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($xRow, $i).Range.Text = $TextToWord.ToString()
									$i++
		                        }
								$xRow++
								$records++
							}
							$selection.EndOf(15) | Out-Null
							$selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.view.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
						}
						"simpletable" {
							Write-Log -Message "writing table type: simpletable" -LogFile $logfile
							$Table = $Null
							$TableRange = $Null
							$TableRange = $doc.Application.Selection.Range
							$Columns = 0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }
							$Table = $doc.Tables.Add($TableRange, $Columns, 2)
							$table.Style = $TableSimpleStyle
							$i = 1;
							Write-Log -Message ("Table: $Columns rows and 2 columns") -LogFile $logfile
							$records = 1
							$y=0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }

								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								$Table.Cell($i, 1).Range.Font.Bold = $true
								$Table.Cell($i, 1).Range.Text = $field.Description
								$Table.Cell($i, 2).Range.Font.Bold = $false
								if ($poshversion -ne 3) { 
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.Rows[0].$($field.FieldName)
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								}
								else {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.$($field.FieldName) 
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								}
								$i++
								$records++
							}
					        $selection.EndOf(15) | Out-Null
					        $selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.View.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
							break
						}
						default {
                            Write-Log -Message "writing table type: default" -LogFile $logfile
							$records = 1
							$y=0
		                    foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
		                        foreach ($field in $HealthCheck.Fields.Field) {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = ($field.Description + " : " + (Get-MessageInformation -MessageID ($row.$($field.FieldName))))
											break ;
										}
										"messagesolution" {
											$TextToWord = ($field.Description + " : " + (Get-MessageSolution -MessageID ($row.$($field.FieldName))))
											break ;
										}												
										default {
											$TextToWord = ($field.Description + " : " + $row.$($field.FieldName))
											break;
										}
									}
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									Write-WordText -WordSelection $selection -Text ($TextToWord.ToString()) -NewLine $true
		                        }
								$selection.TypeParagraph()
								$records++
		                    }
						}
                	}
				}
			}
		}
        if ($bFound -eq $false) {
		    Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
		    Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
		    $selection.TypeParagraph()
		}
	}
}

function Write-WordTableGrid {
    param (
        [parameter(Mandatory=$True)]
        [string] $Caption,
        [parameter(Mandatory=$True)]
        [int] $Rows,
        [parameter(Mandatory=$True)]
        [string[]] $ColumnHeadings
	)
	Write-Log -Message "inserting custom table: $Caption" -LogFile $logfile
    $Selection.TypeText($Caption)
    $Selection.Style = "Heading 1"
    $Selection.TypeParagraph()
    $Cols  = $ColumnHeadings.Length
    $Table = $doc.Tables.Add($Selection.Range, $rows, $cols)
    $Table.Style = "Grid Table 4 - Accent 1"
    for ($col = 1; $col -le $cols; $col++) {
        $Table.Cell(1, $col).Range.Text = $ColumnHeadings[$col-1]
    }
    for ($row = 1; $row -lt $rows; $row++) {
        $Table.Cell($row+1, 1).Range.Text = $row.ToString()
    }
    # set table width to 100%
    $Table.PreferredWidthType = 2
    $Table.PreferredWidth = 100
    # set column widths
    if ($cols -gt 2) {
        $Table.Columns(2).PreferredWidthType = 2
        $Table.Columns(2).PreferredWIdth = 7
    }
    else {
        $Table.Columns.First.PreferredWidthType = 2
        $Table.Columns.First.PreferredWidth = 7
    }
    $Selection.EndOf(15) | Out-Null
	$Selection.MoveDown() | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = 0
	$Selection.EndKey(6, 0) | Out-Null
	$Selection.TypeParagraph()
}

<#
.SYNOPSIS
	Get-CmSiteInstallPath returns [string] path to the base installation
	of System Center Configuration Manager on the site server.
.DESCRIPTION
	Returns the full SCCM installation path using a registry query.
#>
function Get-CmSiteInstallPath {
	Write-Log -Message "getting configmgr installation path" -Logfile $logfile
	try {
		$x = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\setup"
		Write-Output $x.'Installation Directory'
	}
	catch {}
}

function Get-CmHealthCheckFile {
	param (
		[parameter(Mandatory=$True, HelpMessage="XML file source path")]
		[ValidateNotNullOrEmpty()]
		[string] $XmlSource
	)
	Write-Log -Message "-----------------------------------------------------" -LogFile $logfile
	if ($XmlSource.StartsWith('http')) {
        Write-Log -Message "importing xml from remote URI: $XmlSource" -LogFile $logfile
        try {
			[xml]$result = ((New-Object System.Net.WebClient).DownloadString($XmlSource))
        }
        catch {
            Write-Error "Failed to import data from Uri: $XmlSource"
            break
        }
        Write-Log -Message "configuration XML data loaded successfully" -LogFile $logfile
    }
    else {
        Write-Log -Message "importing Configuration xml from local file: $XmlSource"
        if (!(Test-Path -Path $XmlSource)) {
            Write-Warning "File $XmlSource does not exist, no futher action taken"
            break
        }
        else { 
            try {
                [xml]$result = Get-Content ($XmlSource) 
            }
            catch {
                Write-Error "Failed to import data from local file: $XmlSource"
                break
            }
            Write-Log -Message "configuration XML data loaded successfully" -LogFile $logfile
        }
	}
	Write-Output $result
}

function Set-WordFormatting {
	if ($WordVersion -ge "16.0") {
		Write-Log -Message "setting styles for Word 2016" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "15.0") {
		Write-Log -Message "setting styles for Word 2013" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "14.0") {
		Write-Log -Message "setting styles for Word 2010" -LogFile $logfile
		$x1 = "Medium Shading 1 - Accent 1"
		$x2 = "Light Grid - Accent 1"
	}
	Write-Output @($x1, $x2)
}