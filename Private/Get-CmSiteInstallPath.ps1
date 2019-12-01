function Get-CmSiteInstallPath {
	Write-Log -Message "getting configmgr installation path" -Logfile $logfile
	try {
		$x = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\setup"
		Write-Output $x.'Installation Directory'
	}
	catch {}
}
