function Get-MessageSolution {
    [CmdletBinding()]
    param (
		$MessageID
    )
    Write-Verbose "looking up error message id: $MessageID"
	$msg = $MessagesXML.dtsHealthCheck.MessageSolution | Where-Object {$_.MessageId -eq $MessageID}
	if ($null -eq $msg)	{
        Write-Verbose "searching windows update error table"
        $errcodes = Join-Path $(Split-Path (Get-Module "cmhealthcheck").Path) -ChildPath "assets\windows_update_errorcodes.txt"
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
            Write-Output "There is no known possible solution for Message ID $MessageID"
        }
    }
	else {
        Write-Output $msg.Description
    }
}