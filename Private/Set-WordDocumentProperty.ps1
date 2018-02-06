Function Set-WordDocumentProperty {
    param (
		[parameter(Mandatory=$True)]
			$Document,
		[parameter(Mandatory=$True)]
			[string] $Name,
		[parameter(Mandatory=$True)]
			[string] $Value
	)
    Write-Log -Message "info: document property [$Name] set to [$Value]" -LogFile $logfile
    $document.BuiltInDocumentProperties($Name) = $Value
}