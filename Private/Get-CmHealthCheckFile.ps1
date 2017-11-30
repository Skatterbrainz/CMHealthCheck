function Get-CmHealthCheckFile {
	param (
		[parameter(Mandatory=$True, HelpMessage="XML file source path")]
		[ValidateNotNullOrEmpty()]
		[string] $XmlSource
	)
	Write-Log -Message "-----------------------------------------------------" -LogFile $logfile
	if ($XmlSource.StartsWith('http')) {
        Write-Log -Message "importing xml from remote URI: $XmlSource" -LogFile $logfile
        try {
			[xml]$result = ((New-Object System.Net.WebClient).DownloadString($XmlSource))
        }
        catch {
            Write-Error "Failed to import data from Uri: $XmlSource"
            break
        }
        Write-Log -Message "configuration XML data loaded successfully" -LogFile $logfile
    }
    else {
        Write-Log -Message "importing Configuration xml from local file: $XmlSource"
        if (!(Test-Path -Path $XmlSource)) {
            Write-Warning "File $XmlSource does not exist, no futher action taken"
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