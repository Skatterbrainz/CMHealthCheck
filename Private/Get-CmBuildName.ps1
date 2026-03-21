function Get-CmBuildName {
	[CmdletBinding()]
	param (
		[parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[string] $BuildNumber
	)
	Write-Log -Message "(Get-CmBuildName)" -LogFile $logfile
	$moduleRoot = Split-Path -Parent $PSScriptRoot
	$bdatafile  = Join-Path -Path $moduleRoot -ChildPath "Assets/buildnumbers.txt"
	if (!(Test-Path $bdatafile)) {
		Write-Error "$bdatafile could not be found or imported"
		return $null
	}
	$bdata = Get-Content $bdatafile -ErrorAction Stop | Where-Object {
		(-not [string]::IsNullOrWhiteSpace($_)) -and (-not $_.StartsWith(';'))
	}
	foreach ($row in $bdata) {
		$bset = $row -split "=", 2
		if ($bset.Count -lt 2) { continue }
		$bnum = $bset[0]
		if ($bnum -eq $BuildNumber) {
			$result = $bset[1]
			return $result
		}
	}
	return $null
}
