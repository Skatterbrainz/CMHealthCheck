Function Get-ServiceStatus {
    param (
        $LogFile,
        [string] $ServerName,
				[parameter(Mandatory=$True)]
        [string] $ServiceName
    )
    Write-Log -Message "function... Get-ServiceStatus ****" -LogFile $logfile
    Write-Log -Message "servername. $servername" -LogFile $logfile
    Write-Log -Message "service.... $servicename" -LogFile $logfile
    try {
        $service = Get-Service -ComputerName $servername | Where-Object {$_.Name -eq $servicename}
        if ($service -ne $null) { $return = $service.Status }
        else { $return = "ERROR: Not Found" }
        Write-Log -Message "status..... $return" -LogFile $logfile
    }
    catch {
        $return = "ERROR: Unknown"
        $Error.Clear()
    }
    Write-Output $return
}