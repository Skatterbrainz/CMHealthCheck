Function Get-SQLServerConnection {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $SQLServer,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $DBName
	)
	Write-Log -Message "(Get-SQLServerConnection)" -LogFile $logfile
	try {
		$conn = New-Object System.Data.SqlClient.SqlConnection
		$conn.ConnectionString = "Data Source=$SQLServer;Initial Catalog=$DBName;Integrated Security=SSPI;"
		return $conn
	} catch {
		$errorMessage = $_.Exception.Message
		$errorCode = "0x{0:X}" -f $_.Exception.ErrorCode
		Write-Log -Message "The following error happen, no futher action taken" -LogFile $logfile
		Write-Log -Message "Error $errorCode : $errorMessage connecting to $SQLServer" -Severity 3 -LogFile $logfile
		$Error.Clear()
		throw "Error $errorCode : $errorMessage connecting to $SQLServer"
	}
}