function Get-CmCredentials {
	Write-Log -Message "(Get-CmCredentials)" -LogFile $logfile
	try {
		$cred = Get-Credentials
		Write-Log -Message "Trying username: $($cred.Username)" -LogFile $logfile
		Write-Output $cred
	} catch {
		Write-Output $null
	}
}