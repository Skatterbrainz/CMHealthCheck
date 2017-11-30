Function Test-RegistryExist {
    param (
		$ComputerName,
		$LogFile = '' ,
		$KeyName,
		$AccessType = 'LocalMachine'
    )
	Write-Log -Message "Testing registry key from $($ComputerName), $($AccessType), $($KeyName)" -LogFile $logfile
    try {
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
		$RegKey = $Reg.OpenSubKey($KeyName)
		$return = ($RegKey -ne $null)
    }
    catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
    }
    Write-Output $return
}