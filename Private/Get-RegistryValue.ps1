Function Get-RegistryValue {
    param (
        [String] $ComputerName,
        [parameter(Mandatory=$False)]
            [string] $LogFile = '' ,
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [string] $KeyName,
        [string] $KeyValue,
        [parameter(Mandatory=$False)]
            [ValidateSet('LocalMachine','ClassesRoot','CurrentConfig','Users')]
            [string] $AccessType = 'LocalMachine'
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
    Write-Output $return
}