function Write-SqlMemory {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
	)
    Write-Log -Message "function... Write-SqlMemory ****" -LogFile $logfile
    $memData = Get-DbaMaxMemory -SqlInstance $ServerName -ErrorAction SilentlyContinue
    if ($null -eq $memData) { return }
	$Fields = @("TotalMemory","MaxLimit")
    $memDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    $row = $memDetails.NewRow()
    $row.TotalMemory = $memData.Total
    $row.MaxLimit = $memData.MaxValue
    $memDetails.Rows.Add($row)
    , $memDetails | Export-CliXml -Path ($filename)
}
