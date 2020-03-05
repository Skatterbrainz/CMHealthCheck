function Write-DocReportSections {
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '1' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '2' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '3' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '4' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '5' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	if ($detailed -eq $true) {
		Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '5' -Detailed $true -Doc $doc -Selection $selection -LogFile $logfile 
		Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	}
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '6' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
}