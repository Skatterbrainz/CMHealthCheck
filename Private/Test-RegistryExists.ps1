function Test-RegistryExist {
	param (
		[parameter()][string] $ComputerName,
		[parameter()][string] $LogFile = '' ,
		[parameter(Mandatory)][string] $KeyName,
		[parameter()][string] $AccessType = 'LocalMachine'
	)
	Write-Log -Message "(Test-RegistryExists)" -LogFile $logfile
	Write-Log -Message "computer... $ComputerName" -LogFile $logfile
	Write-Log -Message "accesstype. $AccessType" -LogFile $logfile
	Write-Log -Message "keyname.... $KeyName" -LogFile $logfile
	try {
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
		$RegKey = $Reg.OpenSubKey($KeyName)
		$result = ($null -ne $RegKey)
	}
	catch {
		$result = "ERROR: Unknown"
		$Error.Clear()
	}
	Write-Output $result
}