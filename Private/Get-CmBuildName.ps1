function Get-CmBuildName {
	param (
		[parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[string] $BuildNumber
	)
	Write-Log -Message "(Get-CmBuildName)" -LogFile $logfile
	$ModuleData = Get-Module CMHealthCheck
	$ModuleVer  = $ModuleData.Version -join '.'
	$ModulePath = $ModuleData.Path -replace 'CMHealthCheck.psm1', ''
	$bdatafile  = "$ModulePath"+"assets\buildnumbers.txt"
	if (!(Test-Path $bdatafile)) {
		Write-Error "$bdatafile could not be found or imported"
		break
	}
	$bdata = Get-Content $bdatafile
	foreach ($row in $bdata) {
		$bset = $row -split "="
		$bnum = $bset[0]
		if ($bnum -eq $BuildNumber) {
			$result = $bset[1]
			break
		}
	}
	Write-Output $result
}
