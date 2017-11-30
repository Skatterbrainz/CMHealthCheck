Function Set-WordDocumentProperty {
    param (
		$Document,
		$Name,
		$Value
	)
    Write-Log -Message "info: document property [$Name] set to [$Value]" -LogFile $logfile
    $document.BuiltInDocumentProperties($Name) = $Value
}