function Get-CmWmiObject {
	param (
		[parameter()] $Class,
		[parameter()][string] $Filter = '',
		[parameter()][string] $Query = '',
		[parameter()][string] $ComputerName,
		[parameter()][string] $Namespace = "root\cimv2",
		[parameter()][string] $LogFile,
		[parameter()][bool] $ContinueOnError = $false
	)
	Write-Log -Message "(Get-CmWmiObject)" -LogFile $logfile
	if (![string]::IsNullOrEmpty($query)) {
		$class = $Query
	}
	Write-Log -Message "WMI Query: \\$ComputerName\$Namespace, $class, filter: $filter" -LogFile $logfile
	if (![string]::IsNullOrEmpty($query)) {
		$WMIObject = Get-CimInstance -Query $query -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
	elseif (![string]::IsNullOrEmpty($filter)) { 
		$WMIObject = Get-CimInstance -ClassName $class -Filter $filter -Namespace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
	else { 
		$WMIObject = Get-CimInstance -ClassName $class -NameSpace $Namespace -ComputerName $ComputerName -ErrorAction SilentlyContinue 
	}
	if ($null -eq $WMIObject) {
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