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