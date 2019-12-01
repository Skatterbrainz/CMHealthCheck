Function Set-WordDocumentProperty {
    param (
		[parameter(Mandatory)] $Document,
		[parameter(Mandatory)] [string] $Name,
		[parameter(Mandatory)] [string] $Value
	)
    Write-Log -Message "info: document property [$Name] set to [$Value]" -LogFile $logfile
    $document.BuiltInDocumentProperties($Name) = $Value
}