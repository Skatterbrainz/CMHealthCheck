function Get-CmXMLFile {
	param (
		[parameter(Mandatory=$True, HelpMessage="Path to file")]
			[ValidateNotNullOrEmpty()]
			[string] $Path,
		[parameter(Mandatory=$True, HelpMessage="File name")]
			[ValidateNotNullOrEmpty()]
			[string] $FileName
	)
	$xfile = Join-Path -Path $Path -ChildPath $FileName
	Write-Log -Message "(Get-CmXMLFile): $xfile" -LogFile $logfile
	try {
		$result = Import-CliXml -Path $xfile
	} catch {
		Write-Log -Message "File $xfile not found. No further action taken" -LogFile $logfile -Severity 3 -ShowMsg
		Write-Log -Message "Typical reasons are: interrupted data collection, permissions into SQL instance or host OS on remote systems." -LogFile $logfile -Severity 3 -ShowMsg
		Write-Log -Message "Recommend: Review Get-CMHealthCheck.log for more information." -LogFile $logfile -Severity 3 -ShowMsg
		break
	}
	Write-Output $result
}
