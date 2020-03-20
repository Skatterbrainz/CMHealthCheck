function Get-XmlUrlContent {
	param (
		[parameter(Mandatory=$True, HelpMessage="Target URL")]
		[ValidateNotNullOrEmpty()]
		[string] $Url
	)
	Write-Log -Message "(Get-XmlUrlContent): $Url" -LogFile $logfile
	$content = ""
	try {
		[xml]$content = ((New-Object System.Net.WebClient).DownloadString($Url))
	}
	catch {}
	if (![string]::IsNullOrEmpty($content)) {
		$lines = $content -split "`n"
		$result = ""
		for ($i = 1; $i -lt $lines.count; $i++) {
			$result += $lines[$i] + "`n"
		}
	}
	Write-Output $result
}