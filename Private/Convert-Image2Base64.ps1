<#
Don't hate me for not being able to remember who wrote this, but it wasn't me.
#>
function Convert-Image2Base64 {
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipelineByPropertyName=$false,Mandatory=$true,ValueFromPipeline=$True,
		HelpMessage="this is a path to either a file on the web or locally on the network to convert")]
		[string]$Path
	)
	Write-Log -Message "(Convert-Image2Base64): $Path" -LogFile $logfile
	if (($Path -match '\.jpg') -or ($Path -match '\.jpeg') -or ($Path -match '\.png')) {
		Write-Log -Message "importing image from file system" -LogFile $logfile
		if (Test-Path -Path "$Path") {
			if ($PSVersionTable.PSVersion.Major -eq 5) {
				$content = Get-Content $Path -Encoding Byte
			}
			else {
				$content = Get-Content $Path -AsByteStream
			}
			$EncodedImage = [convert]::ToBase64String($content)
		}
		else {
			Write-Log -Message "Image file not found: $path" -LogFile $logFile -Severity 3
			Return $false
		}
	}
	elseif ($Path -match '^http[s]://.*(\.png|\.jpg)$') {
		Write-Log -Message "importing image from URL" -LogFile $logfile
		$ext=$Path.Substring($Path.Length-4)
		$tempfile = "${env:TEMP}\logo31337$ext"
		if (Test-Path $tempfile) {Remove-Item -Path $tempfile -Force}
		try {
			Invoke-WebRequest -Uri $Path -OutFile $tempfile
		}
		catch {
			Write-Log -Message "logo image file not found, returning nothing" -LogFile $logfile
			Return $false
		}
		if ($PSVersionTable.PSVersion.Major -eq 5) {
			$content = Get-Content $tempfile -Encoding Byte
		}
		else {
			$content = Get-Content $tempfile -AsByteStream
		}
		$EncodedImage = [convert]::ToBase64String($content)
	}
	else {
		Write-Log -Message "Path does not match pattern: $path" -LogFile $logfile -Severity 3
		Return $false
	}
	if ($path.EndsWith(".jpg")) {
		$imgtype = "jpg"
	}
	elseif ($path.EndsWith(".png")) {
		$imgtype = "png"
	}
	Write-Log -Message "imagetype = $imgtype" -LogFile $logfile
	"data:image/$imgtype;base64,$EncodedImage"
}