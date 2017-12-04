function Set-WordFormatting {
	Write-Log -Message "function... Set-WordFormatting ****" -LogFile $logfile
	if ($WordVersion -ge "16.0") {
		Write-Log -Message "setting styles for Word 2016" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "15.0") {
		Write-Log -Message "setting styles for Word 2013" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	}
	elseif ($WordVersion -eq "14.0") {
		Write-Log -Message "setting styles for Word 2010" -LogFile $logfile
		$x1 = "Medium Shading 1 - Accent 1"
		$x2 = "Light Grid - Accent 1"
	}
	Write-Output @($x1, $x2)
}