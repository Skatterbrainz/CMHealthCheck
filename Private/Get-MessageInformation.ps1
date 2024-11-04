function Get-MessageInformation {
	[CmdletBinding()]
	param (
		[parameter()][string]$MessageID = ""
	)
	Write-Log -Message "(Get-MessageInformation): MessageID = $MessageID"
	try {
		if (![string]::IsNullOrEmpty($MessageID)) {
			Write-Log -Message "looking up information for error messageID = $MessageID" -LogFile $logfile
			$msg = $MessagesXML.dtsHealthCheck.Message | Where-Object {$_.MessageId -eq $MessageID}
			if ($null -eq $msg) {
				Write-Log -Message "searching windows update error table" -LogFile $logfile
				$errcodes = Join-Path $(Split-Path (Get-Module "cmhealthcheck").Path) -ChildPath "assets\windows_update_errorcodes.csv"
				if (Test-Path $errcodes) {
					Write-Log -Message "importing lookup data" -LogFile $logfile
					$errdata = Import-Csv -Path $errcodes
					$errdet = $($errdata | Where-Object {$_.ErrorCode -eq $MessageID} | Select-Object -ExpandProperty Description).Trim()
					if ([string]::IsNullOrEmpty($errdet)) {
						$errdet = $($errdata | Where-Object {$_.DecErrorCode -eq $MessageID} | Select-Object -ExpandProperty Description).Trim()
						if (![string]::IsNullOrEmpty($errdet)) {
							Write-Output $errdet
						} else {
							Write-Output "There is no known possible solution for Message ID $MessageID"
						}
					} else {
						Write-Output $errdet
					}
				} else {
					Write-Log -Message "file not found: $errcodes" -LogFile $logfile -Severity 3
					Write-Output "Unknown Message ID $MessageID"
				}
			} else {
				Write-Log -Message "reading xml message description" -LogFile $logfile
				Write-Output $msg.Description
			}
		} else {
			throw "MessageID was blank or null"
		}
	} catch {
		Write-Log -Message "Error: $($_.Exception.Message -join ';')" -LogFile $logfile -Severity 3
		Write-Output ""
	}
}
