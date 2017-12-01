Function Write-DiskInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
	)
	Write-Log -Message "function... Write-DiskInfo ****" -LogFile $logfile
    $DiskList = Get-CmWmiObject -Class "Win32_LogicalDisk" -Filter "DriveType = 3" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror
    if ($DiskList -eq $null) { return }
	$Fields=@("DeviceID","Size","FreeSpace","FileSystem")
	$DiskDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($Disk in $DiskList) {
		$row = $DiskDetails.NewRow()
		$row.DeviceID = $Disk.DeviceID
		$row.Size = ([int](($Disk.Size) / 1024 / 1024 / 1024)).ToString()
		$row.FreeSpace = ([int](($Disk.FreeSpace) / 1024 / 1024 / 1024)).ToString()
		$row.FileSystem = $Disk.FileSystem
	    $DiskDetails.Rows.Add($row)
    }
    , $DiskDetails | Export-CliXml -Path ($filename)
}