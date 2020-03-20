function New-CmDataTable {
	param (
		[parameter()][string] $TableName,
		[parameter()][string[]] $Fields
	)
	Write-Log -Message "(New-CmDataTable): $TableName" -LogFile $logfile
	$DataTable = New-Object System.Data.DataTable "$tableName"
	foreach ($field in $fields) {
		$col = New-Object System.Data.DataColumn "$field",([string])
		$DataTable.Columns.Add($col)
	}
	,$DataTable
}
