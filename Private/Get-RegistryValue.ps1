Function Get-RegistryValue {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $KeyName,
		[parameter()][string] $ComputerName,
		[parameter()][string] $LogFile = '' ,
		[parameter()][string] $KeyValue,
		[parameter()][ValidateSet('LocalMachine','ClassesRoot','CurrentConfig','Users')][string] $AccessType = 'LocalMachine'
	)
	Write-Log -Message "Getting registry value from $($ComputerName), $($AccessType), $($keyname), $($keyvalue)" -LogFile $logfile
	try {
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
		$RegKey= $Reg.OpenSubKey($keyname)
		if ($RegKey -ne $null) {
			try { $return = $RegKey.GetValue($keyvalue) }
			catch { $return = $null }
		}
		else { $return = $null }
		Write-Log -Message "Value returned $return" -LogFile $logfile
	}
	catch {
		$return = "ERROR: Unknown"
		$Error.Clear()
	}
	, $return
}