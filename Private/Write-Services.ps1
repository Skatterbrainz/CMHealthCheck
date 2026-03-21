function Write-Services {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ContinueOnError = $true
	)
	Write-Log -Message "(Write-Services)" -LogFile $logfile
	$services = @()
	$svcDetails = $null
	try {
		$services = @(Get-CimInstance -ClassName Win32_Service -ComputerName $ServerName | Select-Object DisplayName,StartName,StartMode,State | Sort-Object DisplayName)
		if ($null -eq $services) { return }
		$Fields = @('DisplayName','StartName','StartMode','State')
		$svcDetails = New-CmDataTable -TableName $tableName -Fields $Fields
		foreach ($service in $services) {
			$row             = $svcDetails.NewRow()
			$row.DisplayName = $service.DisplayName
			$row.StartName   = $service.StartName
			$row.StartMode   = $service.StartMode
			$row.State       = $service.State
			$svcDetails.Rows.Add($row)
		}
	}
	catch {
		Write-Log -Message "Unable to enumerate services on $ServerName: $($_.Exception.Message)" -Severity 3 -LogFile $logfile
		if ($ContinueOnError -eq $false) {
			throw
		}
		return
	}
	Write-Log -Message "enumerated $($services.Count) services" -LogFile $LogFile
	if ($null -ne $svcDetails) {
		, $svcDetails | Export-CliXml -Path ($filename)
	}
}
