function Write-DiscoveryMethods {
    param (
        [string] $FileName,
        [string] $TableName,
        [string] $SiteCode,
        [int] $NumberOfDays,
        $LogFile,
        [string] $ServerName,
        $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-discoverymethods]" -LogFile $logfile
    $query = "select distinct ItemType, Sitenumber, SourceTable from SC_Properties where (ItemType like '%Discover%') order by ItemType"
    $dms = @(Invoke-DbaQuery -SqlInstance $ServerName -Database "CM_$SiteCode" -Query $query -ErrorAction SilentlyContinue)
    if ($null -eq $dms) { return }
    $Fields = @("ItemType", "SiteNumber","SourceTable")
    $dmDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    foreach ($dm in $dms) {
        $row = $dmDetails.NewRow()
        $row.ItemType = $dm.ItemType
        $row.SiteNumber = $dm.SiteNumber
        $row.SourceTable = $dm.SourceTable
        $dmDetails.Rows.Add($row)
    }
    , $dmDetails | Export-CliXml -Path ($filename)
}