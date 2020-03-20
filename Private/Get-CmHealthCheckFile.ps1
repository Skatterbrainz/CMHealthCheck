function Get-CmHealthCheckFile {
	[CmdletBinding()]
	param (
		[parameter(Mandatory, HelpMessage="XML file source path")]
		[ValidateNotNullOrEmpty()]
		[string] $XmlSource
	)
	Write-Log -Message "(Get-CmHealthCheckFile): $XmlSource" -LogFile $logfile
	if ($XmlSource.StartsWith('http')) {
		Write-Log -Message "sourcetype. remote URI" -LogFile $logfile
		try {
			[xml]$result = ((New-Object System.Net.WebClient).DownloadString($XmlSource))
		}
		catch {
			Write-Log -Message "ERROR: failed to import data from $XmlSource" -LogFile $logfile -Severity 3 -ShowMsg
			break
		}
		Write-Log -Message "configuration XML data loaded successfully" -LogFile $logfile
	}
	else {
		Write-Log -Message "sourcetype. localfile" -LogFile $logfile
		Write-Log -Message "filename... $XmlSource" -LogFile $logfile
		if (!(Test-Path -Path $XmlSource)) {
			Write-Log -Message "ERROR: $XmlSource does not exist." -LogFile $logfile -Severity 3 -ShowMsg
			break
		}
		else { 
			try {
				[xml]$result = Get-Content ($XmlSource) 
			}
			catch {
				Write-Error "Failed to import data from local file: $XmlSource"
				break
			}
			Write-Log -Message "configuration XML data loaded successfully" -LogFile $logfile
		}
	}
	Write-Output $result
}