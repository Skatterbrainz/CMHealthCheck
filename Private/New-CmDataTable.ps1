function New-CmDataTable {
    param (
		[string] $TableName,
	    [String[]] $Fields
    )
    Write-Log -Message "[function: New-CmDataTable]" -LogFile $logfile
	$DataTable = New-Object System.Data.DataTable "$tableName"
	foreach ($field in $fields) {
		$col = New-Object System.Data.DataColumn "$field",([string])
		$DataTable.Columns.Add($col)
	}
	,$DataTable
}
