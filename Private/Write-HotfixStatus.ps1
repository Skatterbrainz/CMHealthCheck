Function Write-HotfixStatus {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ContinueOnError = $true
	)
	Write-Log -Message "[function: write-hotfixstatus]" -LogFile $logfile
	try {
		$Session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $ServerName))
		$Searcher = $Session.CreateUpdateSearcher()
		$historyCount = $Searcher.GetTotalHistoryCount()
		$return = $Searcher.QueryHistory(0, $historyCount) 
		Write-Log -Message "  Hotfix count: $HistoryCount" -LogFile $logfile
	}
	catch {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		Write-Log -Message "  The following error happen" -LogFile $logfile
		Write-Log -Message "  Error $errorCode : $errorMessage connecting to $ServerName" -LogFile $logfile
		$Error.Clear()
		return
	}
	$Fields = @("Title", "Date")
	$HotfixTable = New-CmDataTable -tablename $tableName -fields $Fields
	foreach ($hotfix in $return) {
		if (![string]::IsNullOrEmpty($hotfix.Title)) {
			$row       = $HotfixTable.NewRow()
			$row.Title = $hotfix.Title
			$row.Date  = $hotfix.Date
			$HotfixTable.Rows.Add($row)
		}
	}
	, $HotfixTable | Export-CliXml -Path ($filename)
}