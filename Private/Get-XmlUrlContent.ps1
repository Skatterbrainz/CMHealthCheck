function Get-XmlUrlContent {
	param (
		[parameter(Mandatory=$True, HelpMessage="Target URL")]
		[ValidateNotNullOrEmpty()]
		[string] $Url
	)
	Write-Log -Message "(Get-XmlUrlContent): $Url" -LogFile $logfile
	$content = ""
	$result = ""
	try {
		[xml]$content = ((New-Object System.Net.WebClient).DownloadString($Url))
	}
	catch {
		Write-Log -Message "Unable to download XML from $Url : $($_.Exception.Message)" -Severity 3 -LogFile $logfile
		Write-Output $result
		return
	}
	if (![string]::IsNullOrEmpty($content)) {
		$lines = $content -split "`n"
		for ($i = 1; $i -lt $lines.count; $i++) {
			$result += $lines[$i] + "`n"
		}
	}
	Write-Output $result
}