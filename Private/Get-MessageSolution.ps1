function Get-MessageSolution {
	[CmdletBinding()]
	param (
		[parameter()][string]$MessageID = ""
	)
	Write-Log -Message "(Get-MessageSolution): MessageID = $MessageID"
	try {
		if (![string]::IsNullOrEmpty($MessageID)) {
			Write-Log -Message "looking up solution for error message id: $MessageID" -LogFile $logfile
			$msg = $MessagesXML.dtsHealthCheck.MessageSolution | Where-Object {$_.MessageId -eq $MessageID}
			if ([string]::IsNullOrEmpty($msg)) {
				Write-Log -Message "searching windows update error solutions table" -LogFile $logfile
				$errcodes = Join-Path $(Split-Path (Get-Module "cmhealthcheck").Path) -ChildPath "assets\windows_update_errorcodes.csv"
				if (Test-Path $errcodes) {
					Write-Log -Message "importing: $errcodes" -LogFile $logfile
					$errdata = Import-Csv -Path $errcodes
					if (![string]::IsNullOrEmpty($errdata)) {
						Write-Log -Message "imported $($errdata.Count) rows from file" -LogFile $logfile
						$errdet = $($errdata | Where-Object {$_.ErrorCode -eq $MessageID} | Select-Object -ExpandProperty Description).Trim()
						if ([string]::IsNullOrEmpty($errdet)) {
							Write-Log -Message "standard details not found. searching decimal error information" -LogFile $logfile
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
						Write-Log -Message "failed to import $errcodes" -LogFile $logfile
						Write-Output ""
					}
				}
				else {
					Write-Warning "missing file: $errcodes"
					Write-Output "There is no known possible solution for Message ID $MessageID"
				}
			}
			else {
				Write-Output $msg.Description
			}
		}
		else {
			Write-Log -Message "MessageID was blank or null" -LogFile $logfile
			Write-Output ""
		}
	}
	catch {
		Write-Log -Message $_.Exception.Message -LogFile $logfile -Severity 3
	}
}