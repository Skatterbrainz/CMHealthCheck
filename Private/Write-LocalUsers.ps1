function Write-LocalUsers {
	param (
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $Filename,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $TableName,
		[parameter(Mandatory)][string] $SiteCode,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $ServerName,
		[parameter()][string] $LogFile,
		[parameter()][bool] $ContinueOnError
	)
	Write-Log -Message "(Write-LocalGroups)" -LogFile $logfile
	Write-Log -Message "filename... $filename" -LogFile $LogFile
	Write-Log -Message "server..... $ServerName" -LogFile $LogFile
	$ServerShortName = ($ServerName -split '\.')[0]
	try {
		$users = @(Get-CimInstance -ClassName "Win32_UserAccount" -ComputerName $ServerName -Filter "Domain='$ServerShortName'" -ErrorAction Stop | 
			Select-Object Name,FullName,Description,AccountType,AccountExpires,PasswordChangeable,PasswordRequired,SID,LockOut |
				Sort-Object Name)
	} catch {
		Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate local security groups'
		return
	}
	if ($null -eq $users) { return }
	$Fields = @('Name','FullName','Description','AccountType','AccountExpires','PasswordChangeable','PasswordRequired','SID','LockOut')
	$userDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($user in $users) {
		$row = $userDetails.NewRow()
		$row.Name        = $user.Name
		$row.FullName    = $user.FullName
		$row.Description = $user.Description
		$row.AccountType = $user.AccountType
		$row.AccountExpires = $user.AccountExpires
		$row.PasswordChangeable = $user.PasswordChangeable
		$row.PasswordRequired   = $user.PasswordRequired
		$row.SID = $user.SID
		$row.LockOut = $user.LockOut
		$userDetails.Rows.Add($row)
	}
	Write-Log -Message "enumerated $($users.Count) users" -LogFile $LogFile
	, $userDetails | Export-CliXml -Path ($filename)
}