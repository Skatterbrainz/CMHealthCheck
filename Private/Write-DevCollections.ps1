function Write-DevCollections {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true
	)
	Write-Log -Message "(Write-DevCollections)" -LogFile $logfile
	$query = "select Name, CollectionID, Comment, MemberCount from v_Collection where CollectionType = 2 order by Name"
	$colls = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
	if ($null -eq $colls) { return }
	$Fields = @("Name","CollectionID","Comment","MemberCount")
	$collDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($coll in $colls) {
		$row              = $collDetails.NewRow()
		$row.Name         = $coll.Name
		$row.CollectionID = $coll.CollectionID
		$row.Comment      = $coll.Comment
		$row.MemberCount  = [int]($coll.MemberCount)
		$collDetails.Rows.Add($row)
	}
	, $collDetails | Export-CliXml -Path ($filename)
}