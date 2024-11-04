Function Get-ServiceStatus {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $ServerName,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $ServiceName,
		[parameter()]$LogFile
	)
	Write-Log -Message "(Get-ServiceStatus): $ServiceName on $ServerName" -LogFile $logfile
	try {
		$service = Get-CimInstance -ClassName "Win32_Service" -ComputerName $ServerName | 
			Where-Object {$_.Name -eq $ServiceName} | Select-Object -ExpandProperty "State"
		if ($null -ne $service) { $return = $service }
		else { $return = "ERROR: Service $ServiceName Not Found" }
		Write-Log -Message "status..... $return" -LogFile $logfile
	} catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
	}
	, $return
}