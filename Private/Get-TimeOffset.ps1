function Get-TimeOffset {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][datetime] $StartTime
	)
	Write-Log -Message "(Get-TimeOffset)" -LogFile $logfile
	$secs = ((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds).ToString()
	$ts   = [timespan]::FromSeconds($secs)
	Write-Output $ts.ToString("hh\:mm\:ss")
}