function Write-Boundaries {
	param (
		[parameter()][string] $FileName,
		[parameter()][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ContinueOnError = $true
	)
	Write-Log -Message "[function: write-boundaries]" -LogFile $logfile
	$query = "select distinct 
		vSMS_Boundary.DisplayName, 
		vSMS_Boundary.BoundaryID, 
		vSMS_Boundary.Value as BValue, 
		case 
			when BoundaryType = 0 then 'IP Subnet' 
			when BoundaryType = 1 then 'Active Directory Site' 
			when BoundaryType = 2 then 'IPv6 Prefix' 
			when BoundaryType = 3 then 'IP Address Range' 
			else 'UnKnown' end as BoundaryType, 
		case 
			when BoundaryFlags = 0 then 'Fast' 
			when BoundaryFlags = 1 then 'Slow' 
			end as BoundaryFlags, 
		vSMS_BoundaryGroupMembers.GroupID, 
		vSMS_BoundaryGroup.Name as BGName
		from 
			vSMS_Boundary INNER JOIN
			vSMS_BoundaryGroupMembers ON 
			vSMS_Boundary.BoundaryID = vSMS_BoundaryGroupMembers.BoundaryID 
			inner join
			vSMS_BoundaryGroup ON 
			vSMS_BoundaryGroupMembers.GroupID = vSMS_BoundaryGroup.GroupID
		order by DisplayName"
	
	$blist = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
	if ($null -eq $blist) { return }
	$Fields = @("DisplayName", "BoundaryID", "BValue", "BoundaryType", "BoundaryFlags", "BGName")
	$bDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($b in $blist) {
		$row               = $bDetails.NewRow()
		$row.DisplayName   = $b.DisplayName
		$row.BoundaryID    = $b.BoundaryID
		$row.BValue        = $b.BValue
		$row.BoundaryType  = $b.BoundaryType
		$row.BoundaryFlags = $b.BoundaryFlags
		$row.BGName        = $b.BGName
		$bDetails.Rows.Add($row)
	}
	, $bDetails | Export-CliXml -Path ($filename)
}