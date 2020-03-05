function Set-WordTOC {
	Write-Log -Message "inserting table of contents" -LogFile $logfile
	$toc = $BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.Insert($selection.Range,$True) | Out-Null
}
