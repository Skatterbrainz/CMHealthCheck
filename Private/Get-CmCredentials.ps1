function Get-CmCredentials {
	Write-Log -Message "(Get-CmCredentials)" -LogFile $logfile
	try {
		$cred = Get-Credential
		Write-Log -Message "Trying username: $($cred.Username)" -LogFile $logfile
		Write-Output $cred
	} catch {
		Write-Log -Message "Unable to get credentials: $($_.Exception.Message)" -Severity 3 -LogFile $logfile
		Write-Output $null
	}
}