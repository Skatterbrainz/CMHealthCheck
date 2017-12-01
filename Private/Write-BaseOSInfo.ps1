Function Write-BaseOSInfo {
    param (
	    [string] $FileName,
		[string] $TableName,
		[parameter(Mandatory=$True)]
	    [string] $SiteCode,
	    [int] $NumberOfDays,
	    [string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
	)
	Write-Log -Message "function... Write-BaseOsInfo ****" -LogFile $logfile
    $WMIOS = Get-CmWmiObject -Class "win32_operatingsystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($WMIOS -eq $null) { return }	
    $WMICS = Get-CmWmiObject -Class "win32_computersystem" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
	$WMIProcessor = Get-CmWmiObject -Class "Win32_processor" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    $WMITimeZone  = Get-CmWmiObject -Class "Win32_TimeZone" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
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