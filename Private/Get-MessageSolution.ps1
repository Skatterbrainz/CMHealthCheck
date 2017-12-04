function Get-MessageSolution {
    param (
		$MessageID
	)
	$msg = $MessagesXML.dtsHealthCheck.MessageSolution | Where-Object {$_.MessageId -eq $MessageID}
	if ($msg -eq $null)	{ 
        Write-Output "There is no known possible solution for Message ID $MessageID" 
    }
	else { 
        Write-Output $msg.Description 
    }
}