function Get-MessageInformation {
    [CmdletBinding()]
    param (
		[parameter()][string]$MessageID = ""
	)
	if (![string]::IsNullOrEmpty($MessageID)) {
		Write-Log "looking up information for error messageID = $MessageID" -Log $logfile
		$msg = $MessagesXML.dtsHealthCheck.Message | Where-Object {$_.MessageId -eq $MessageID}
		if ($null -eq $msg) {
			Write-Log "searching windows update error table" -Log $logfile
			$errcodes = Join-Path $(Split-Path (Get-Module "cmhealthcheck").Path) -ChildPath "assets\windows_update_errorcodes.csv"
			if (Test-Path $errcodes) {
				$errdata = Import-Csv -Path $errcodes
				$errdet = $($errdata | Where-Object {$_.ErrorCode -eq $MessageID} | Select-Object -ExpandProperty Description).Trim()
				if ([string]::IsNullOrEmpty($errdet)) {
					$errdet = $($errdata | Where-Object {$_.DecErrorCode -eq $MessageID} | Select-Object -ExpandProperty Description).Trim()
					if (![string]::IsNullOrEmpty($errdet)) {
						Write-Output $errdet
					}
					else {
						Write-Output "There is no known possible solution for Message ID $MessageID"
					}
				}
				else {
					Write-Output $errdet
				}
			}
			else {
				Write-Warning "missing file: $errcodes"
				Write-Output "Unknown Message ID $MessageID"
			}
		}
		else {
			Write-Output $msg.Description
		}
	}
	else {
		Write-Log "MessageID was blank or null" -Log $logfile -Severity 3
		Write-Output ""
	}
}
