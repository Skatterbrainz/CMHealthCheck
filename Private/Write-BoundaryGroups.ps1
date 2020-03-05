function Write-BoundaryGroups {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ContinueOnError = $true
	)
	Write-Log -Message "[function: write-cmboundarygroups]" -LogFile $logfile
	$query = "select distinct Name, GroupID, Description, Flags, DefaultSiteCode as SiteCode, CreatedOn, MemberCount as Boundaries, SiteSystemCount as SiteSystems FROM vSMS_BoundaryGroup order by Name"
	$bgs = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
	if ($null -eq $bgs) { return }
	$Fields = @("Name","GroupID","Description","Flags","SiteCode","CreatedOn","Boundaries","SiteSystems")
	$bgDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($bg in $bgs) {
		$row             = $bgDetails.NewRow()
		$row.Name        = $bg.Name
		$row.GroupID     = $bg.GroupID
		$row.Description = $bg.Description
		$row.Flags       = $bg.Flags
		$row.SiteCode    = $bg.SiteCode
		$row.CreatedOn   = $bg.CreatedOn
		$row.Boundaries  = $bg.Boundaries
		$row.SiteSystems = $bg.SiteSystems
		$bgDetails.Rows.Add($row)
	}
	, $bgDetails | Export-CliXml -Path ($filename)
}