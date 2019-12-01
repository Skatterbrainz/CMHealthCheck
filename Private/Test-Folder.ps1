Function Test-Folder {
    param (
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][String] $Path,
        [parameter()][bool] $Create = $true
    )
    if (Test-Path -Path $Path) {
		return $true
	}
    elseif ($Create -eq $true) {
        try {
            New-Item ($Path) -Type Directory -Force | Out-Null
            Write-Output $true
        }
        catch {
            Write-Output $false
        }
    }
    else {
		Write-Output $false
	}
}