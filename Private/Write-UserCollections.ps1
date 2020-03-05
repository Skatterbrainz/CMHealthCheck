function Write-UserCollections {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true
	)
	Write-Log -Message "function... Write-UserCollections ****" -LogFile $logfile
	try {
		$query = "select Name, CollectionID, Comment, MemberCount from v_Collection where CollectionType = 1 order by Name"
		$colls = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query)
		if ($null -eq $colls) { return }
		$Fields = @("Name","CollectionID","Comment","MemberCount")
		$collDetails = New-CmDataTable -TableName $tableName -Fields $Fields
		foreach ($coll in $colls) {
			$row      = $collDetails.NewRow()
			$row.Name = $coll.Name
			$row.CollectionID = $coll.CollectionID
			$row.Comment      = $coll.Comment
			$row.MemberCount  = [int]($coll.MemberCount)
			$collDetails.Rows.Add($row)
		}
		, $collDetails | Export-CliXml -Path ($filename)
	}
	catch {
		Write-Log -Message "$($_.Exception.Message)" -Category Error -Severity 2
	}
}