function Test-Admin {
	$identity  = [System.Security.Principal.WindowsIdentity]::GetCurrent() 
	$principal = New-Object System.Security.Principal.WindowsPrincipal($identity) 
	$admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator 
	$principal.IsInRole($admin) 
}

Function Test-Folder {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][String] $Path,
		[parameter()][bool] $Create = $true
	)
	if (Test-Path -Path $Path) {
		return $true
	} elseif ($Create -eq $true) {
		try {
			New-Item ($Path) -Type Directory -Force | Out-Null
			Write-Output $true
		} catch {
			Write-Output $false
		}
	} else {
		Write-Output $false
	}
}

function Test-Numeric ($x) {
	($x -match '^\d+$')
}

function Test-Powershell64bit {
	Write-Output ([IntPtr]::size -eq 8)
}

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
	} catch {
		$result = "ERROR: Unknown"
		$Error.Clear()
	}
	Write-Output $result
}