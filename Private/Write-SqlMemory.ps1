function Write-SqlMemory {
    param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true
	)
    Write-Log -Message "function... Write-SqlMemory ****" -LogFile $logfile
    try {
        $memData = Get-DbaMaxMemory -SqlInstance $ServerName
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
    catch {
        Write-Log -Message "$($_.Exception.Message)" -Category Error -Severity 2
    }
}
