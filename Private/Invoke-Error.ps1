function Invoke-Error {
	param (
		[parameter()][string] $Message = ""
	)
	$resolvedMessage = $Message
	if ([string]::IsNullOrEmpty($resolvedMessage) -and ($null -ne $Error[0])) {
		$resolvedMessage = $Error[0].Exception.Message
	}
	if ([string]::IsNullOrEmpty($resolvedMessage)) {
		$resolvedMessage = "An unspecified error occurred"
	}
	Write-Log -Message $resolvedMessage -Severity 3 -LogFile $logfile -ShowMsg
	Stop-Transcript -ErrorAction SilentlyContinue
}