Function Get-ServiceStatus {
    param (
        [parameter(Mandatory=$False)]
            $LogFile,
        [parameter(Mandatory=$True)]
            [string] $ServerName,
		[parameter(Mandatory=$True)]
            [string] $ServiceName
    )
    Write-Log -Message "function... Get-ServiceStatus ****" -LogFile $logfile
    Write-Log -Message "servername. $ServerName" -LogFile $logfile
    Write-Log -Message "service.... $ServiceName" -LogFile $logfile
    try {
        $service = Get-Service -ComputerName $ServerName | Where-Object {$_.Name -eq $ServiceName}
        if ($service -ne $null) { $return = $service.Status }
        else { $return = "ERROR: Not Found" }
        Write-Log -Message "status..... $return" -LogFile $logfile
    }
    catch {
        $return = "ERROR: Unknown"
        $Error.Clear()
    }
    , $return
}