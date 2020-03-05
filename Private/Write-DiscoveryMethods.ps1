function Write-DiscoveryMethods {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ContinueOnError = $true
	)
	Write-Log -Message "[function: write-discoverymethods]" -LogFile $logfile
	$query = "select distinct ItemType,ID,Sitenumber,[Name],Value1,Value2,Value3,SourceTable FROM SC_Properties WHERE (ItemType like '%discover%') ORDER BY ItemType, Name"
	$dms = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
	if ($null -eq $dms) { return }
	$Fields = @("ItemType", "SiteNumber","SourceTable","Name","Value1","Value2","Value3")
	$dmDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($dm in $dms) {
		$row             = $dmDetails.NewRow()
		$row.ItemType    = $dm.ItemType
		$row.SiteNumber  = $dm.SiteNumber
		$row.SourceTable = $dm.SourceTable
		$row.Name        = $dm.Name
		$row.Value1      = $dm.Value1
		$row.Value2      = $dm.Value2
		$row.Value3      = $dm.Value3
		$dmDetails.Rows.Add($row)
	}
	, $dmDetails | Export-CliXml -Path ($filename)
}