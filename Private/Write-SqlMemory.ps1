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
    $memPct  = [math]::Round([int]($memData.MaxValue) / [int]($memData.Total), 2)
    if ($null -eq $memData) { return }
	$Fields = @("TotalMemory","MaxLimit","Pct")
    $memDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    $row = $memDetails.NewRow()
    $row.TotalMemory = $memData.Total
    $row.MaxLimit = $memData.MaxValue
    $row.Pct = $memPct
    $memDetails.Rows.Add($row)
    , $memDetails | Export-CliXml -Path ($filename)
}
