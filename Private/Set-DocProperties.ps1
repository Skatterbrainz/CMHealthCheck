function Set-DocProperties {
	Write-Log -Message "(Set-DocProperties)" -LogFile $logfile
	if ($bAutoProps -eq $True) {
		Write-Log -Message "setting document properties" -LogFile $logfile
		$doc.BuiltInDocumentProperties("Title")    = "System Center Configuration Manager HealthCheck"
		$doc.BuiltInDocumentProperties("Subject")  = "Prepared for $CustomerName"
		$doc.BuiltInDocumentProperties("Author")   = $AuthorName
		$doc.BuiltInDocumentProperties("Company")  = $CopyrightName
		$doc.BuiltInDocumentProperties("Category") = "REPORTS"
		$doc.BuiltInDocumentProperties("Keywords") = "sccm,healthcheck,systemcenter,configmgr,$CustomerName"
	}
}