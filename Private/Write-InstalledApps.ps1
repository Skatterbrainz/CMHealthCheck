function Write-InstalledApps {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $Filename,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $TableName,
		[parameter(Mandatory)][string] $SiteCode,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $ServerName,
		[parameter()][string] $LogFile,
		[parameter()][bool] $ContinueOnError
	)
	Write-Log -Message "function... Write-InstalledApps ****" -LogFile $logfile
	Write-Log -Message "filename... $filename" -LogFile $LogFile
	Write-Log -Message "server..... $ServerName" -LogFile $LogFile
	try {
		$Apps = @(Get-CimInstance -ClassName "Win32_Product" -ComputerName $ServerName -ErrorAction Stop | Sort-Object Name)
	}
	catch {
		if ($ContinueOnError -eq $True) {
			Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate software' -LogFile $LogFile
		}
		else {
			Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate software' -Severity 3 -LogFile $LogFile
			return
		}
	}
	if ($null -eq $Apps) {
		Write-Log -Message "found NO installed applications (aborting)" -LogFile $LogFile
		return
	}
	Write-Log -Message "found $($Apps.Count) installed applications" -LogFile $LogFile
	$Fields = @("Name","Version","Vendor")
	$AppDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($app in $Apps) {
		$appname  = $app.Name
		$appver   = $app.Version 
		$appven   = $app.Vendor 
		$row      = $AppDetails.NewRow()
		$row.Name = $appname
		$row.Version = $appver
		$row.Vendor  = $appven
		$AppDetails.Rows.Add($row)
	}
	, $AppDetails | Export-CliXml -Path ($filename)
}