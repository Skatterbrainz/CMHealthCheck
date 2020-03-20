function Write-NetworkInfo {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true
	)
	Write-Log -Message "(Write-NetworkInfo)" -LogFile $logfile
	$IPDetails = Get-CmWmiObject -Class "Win32_NetworkAdapterConfiguration" -Filter "IPEnabled = true" -ComputerName $ServerName -LogFile $logfile -ContinueOnError $ContinueOnError
	if ($null -eq $IPDetails) { return }
	$Fields = @("IPAddress","DefaultIPGateway","IPSubnet","MACAddress","DHCPEnabled")
	$NetworkInfoTable = New-CmDataTable -TableName $TableName -Fields $Fields
	foreach ($IPAddress in $IPDetails) {
		$row                  = $NetworkInfoTable.NewRow()
		$row.IPAddress        = ($IPAddress.IPAddress -join ", ")
		$row.DefaultIPGateway = ($IPAddress.DefaultIPGateway -join ", ")
		$row.IPSubnet         = ($IPAddress.IPSubnet -join ", ")
		$row.MACAddress       = $IPAddress.MACAddress
		if ($IPAddress.DHCPEnable -eq $true) { 
			$row.DHCPEnabled = "TRUE" 
		}
		else { 
			$row.DHCPEnabled = "FALSE" 
		}
		$NetworkInfoTable.Rows.Add($row)
	}
	, $NetworkInfoTable | Export-CliXml -Path ($FileName)
}