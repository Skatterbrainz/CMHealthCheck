function Write-MPConnectivity {
	param (
		[parameter()][string] $FileName,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $TableName,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $Type = 'mplist'
	)
	Write-Log -Message "(Write-MpConnectivity)" -LogFile $logfile
 	$Fields = @("ServerName", "HTTPReturn")
	$MPConnectivityTable = New-CmDataTable -TableName $tableName -Fields $Fields
	$MPList = Get-CmWmiObject -Query "select * from SMS_SCI_SysResUse where SiteCode = '$SiteCode' and RoleName = 'SMS Management Point'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
	foreach ($MPInformation in $MPList) {
		$SSLState = ($MPInformation.Props | Where-Object {$_.PropertyName -eq "SslState"}).Value
		$mpname = $MPInformation.NetworkOSPath -replace '\\', ''
		if ($SSLState -eq 0) {
			$protocol = 'http'
			$port = $HTTPport 
		} else {
			$protocol = 'https'
			$port = $HTTPSport 
		}
		$web = New-Object -ComObject msxml2.xmlhttp
		$url = $protocol + '://' + $mpname + ':' + $port + '/sms_mp/.sms_aut?' + $type
		if ($healthcheckdebug) { Write-Verbose ("URL Connection: $url") }
		$row = $MPConnectivityTable.NewRow()
		$row.ServerName = $mpname
		try {   
			$web.open('GET', $url, $false)
			$web.send()
			$row.HTTPReturn = $web.status
		} catch {
			$row.HTTPReturn = "313 - Unable to connect to host"
			$Error.Clear()
		}
		Write-Log -Message "status..... $($web.status)" -LogFile $logfile
		$MPConnectivityTable.Rows.Add($row)
	}
	, $MPConnectivityTable | Export-CliXml -Path ($filename)
}