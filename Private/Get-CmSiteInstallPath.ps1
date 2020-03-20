function Get-CmSiteInstallPath {
	Write-Log -Message "(Get-CmSiteInstallPath)" -Logfile $logfile
	try {
		$x = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\setup"
		Write-Output $x.'Installation Directory'
	}
	catch {}
}
