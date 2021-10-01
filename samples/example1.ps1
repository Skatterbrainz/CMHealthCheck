#requires -RunAsAdministrator
[CmdletBinding()]
param()

if (not(Get-Module cmHealthCheck -ListAvailable)) {
	try {
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
		Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
	}
	catch {
		Write-Output "error: $($_.Exception.Message -join ';')"
	}
}

$params = @{
	SmsProvider = "cm01.contoso.local"
	CustomerName = "Contoso"
	Author = "Your Name"
	CopyrightName = "Contoso Corporation"
	NoHotFix = $True
	Detailed = $True
	Overwrite = $True
}

Invoke-CMHealthCheck @params