function Get-CmWmiObject {
    param (
		$Class,
		[parameter(Mandatory=$False)]
			[string] $Filter = '',
		[parameter(Mandatory=$False)]
			[string] $Query = '',
		[parameter(Mandatory=$False)]
			$ComputerName,
		[parameter(Mandatory=$False)]
			[string] $Namespace = "root\cimv2",
		[parameter(Mandatory=$False)]
			[string] $LogFile,
		[parameter(Mandatory=$False)]
			[bool] $ContinueOnError = $false
    )
    if ($query -ne '') { $class = $query }
	Write-Log -Message "WMI Query: \\$ComputerName\$Namespace, $class, filter: $filter" -LogFile $logfile
    if ($query -ne '') { 
		$WMIObject = Get-WmiObject -Query $query -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
    elseif ($filter -ne '') { 
		$WMIObject = Get-WmiObject -Class $class -Filter $filter -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
    else { 
		$WMIObject = Get-WmiObject -Class $class -NameSpace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
	if ($WMIObject -eq $null) {
		Write-Log -Message "WMI Query returned 0) records" -LogFile $logfile
	}
	else {
		$i = 1
		foreach ($obj in $wmiobj) { i++ }
		Write-Log -Message "WMI Query returned $($i) records" -LogFile $logfile
	}
    if ($Error.Count -ne 0) {
		$errorMessage = $Error[0].Exception.Message
		$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
		if ($ContinueOnError -eq $false) {
            Write-Log -Message "The following error occurred, no futher action taken" -Severity 3 -Logfile $logfile
        }
		else { 
            Write-Error "The following error occurred"
        }
		Write-Log -Message "ERROR $errorCode : $errorMessage connecting to $ComputerName" -LogFile $logfile
		$Error.Clear()
		if ($ContinueOnError -eq $false) { 
            Throw "ERROR $errorCode : $errorMessage connecting to $ComputerName" 
        }
    }
    Write-Output $WMIObject
}