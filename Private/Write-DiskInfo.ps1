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
    $DiskList = @(Get-CmWmiObject -Class "Win32_LogicalDisk" -Filter "DriveType = 3" -ComputerName $servername -LogFile $logfile -ContinueOnError $continueonerror)
    if ($DiskList -eq $null) { return }
	$Fields = @("DeviceID","VolumeName","FileSystem","Size","FreeSpace","Used","PctUsed")
	$DiskDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($Disk in $DiskList) {
		$used = ($Disk.Size - $Disk.FreeSpace)
		$pct  = ($used / $Disk.Size)
		$row  = $DiskDetails.NewRow()
		$row.DeviceID   = $Disk.DeviceID
		$row.VolumeName = $Disk.VolumeName
		$row.FileSystem = $Disk.FileSystem
		$row.Size       = ([int](($Disk.Size) / 1024 / 1024 / 1024)).ToString()
		$row.FreeSpace  = ([int](($Disk.FreeSpace) / 1024 / 1024 / 1024)).ToString()
		$row.Used       = ([int]($used / 1024 / 1024 / 1024)).ToString()
		$row.PctUsed    = [math]::Round($pct,2)
	    $DiskDetails.Rows.Add($row)
    }
    , $DiskDetails | Export-CliXml -Path ($filename)
}