function Get-MessageInformation {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.Message | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null) {
        Write-Output "Unknown Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}
