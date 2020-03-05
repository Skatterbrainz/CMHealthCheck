function Show-CMHCInfo {
	Write-Log -Message "==========" -LogFile $logfile -ShowMsg
	Write-Log -Message "Script Version...: $ModuleVer" -LogFile $logfile -ShowMsg
	Write-Log -Message "Windows Version..: $osversion" -LogFile $logfile
	Write-Log -Message "environment......: Running Powershell version: $poshversion" -LogFile $logfile
	Write-Log -Message "environment......: Running Powershell 64 bits" -LogFile $logfile
	Write-Log -Message "Report Folder....: $reportFolder" -LogFile $logfile
	Write-Log -Message "Detailed Report..: $detailed" -LogFile $logfile
	Write-Log -Message "Current Folder..: $($PWD.Path)" -LogFile $logfile
	Write-Log -Message "Log Folder......: $logFolder" -LogFile $logfile
	Write-Log -Message "Log File........: $logfile" -LogFile $logfile
	Write-Log -Message "control file....: $Healthcheckfilename" -LogFile $logfile
	Write-Log -Message "message file....: $MessagesFilename" -LogFile $logfile
}