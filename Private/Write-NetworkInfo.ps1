function Write-NetworkInfo {
    param (
		[string] $FileName,
		[string] $TableName,
		[string] $SiteCode,
		[int] $NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		[bool] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-networkinfo]" -LogFile $logfile
    $IPDetails = Get-CmWmiObject -Class "Win32_NetworkAdapterConfiguration" -Filter "IPEnabled = true" -ComputerName $ServerName -LogFile $logfile -ContinueOnError $ContinueOnError
    if ($IPDetails -eq $null) { return }
	$Fields = @("IPAddress","DefaultIPGateway","IPSubnet","MACAddress","DHCPEnabled")
	$NetworkInfoTable = New-CmDataTable -TableName $TableName -Fields $Fields
	foreach ($IPAddress in $IPDetails) {
		$row = $NetworkInfoTable.NewRow()
		$row.IPAddress = ($IPAddress.IPAddress -join ", ")
		$row.DefaultIPGateway = ($IPAddress.DefaultIPGateway -join ", ")
		$row.IPSubnet = ($IPAddress.IPSubnet -join ", ")
		$row.MACAddress = $IPAddress.MACAddress
		if ($IPAddress.DHCPEnable -eq $true) { $row.DHCPEnabled = "TRUE" } else { $row.DHCPEnabled = "FALSE" }
	    $NetworkInfoTable.Rows.Add($row)
    }
    , $NetworkInfoTable | Export-CliXml -Path ($FileName)
}