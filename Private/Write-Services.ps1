function Write-Services {
    param (
        [string] $FileName,
        [string] $TableName,
        [string] $SiteCode,
        [int] $NumberOfDays,
        $LogFile,
        [string] $ServerName,
        $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-services]" -LogFile $logfile
    try {
        $services = @(Get-WmiObject -Class Win32_Service -ComputerName $ServerName | Select-Object DisplayName,StartName,StartMode,State | Sort-Object DisplayName)
        if ($null -eq $services) { return }
        $Fields = @('DisplayName','StartName','StartMode','State')
        $svcDetails = New-CmDataTable -TableName $tableName -Fields $Fields
        foreach ($service in $services) {
            $row = $svcDetails.NewRow()
            $row.DisplayName = $service.DisplayName
            $row.StartName = $service.StartName
            $row.StartMode = $service.StartMode
            $row.State = $service.State
            $svcDetails.Rows.Add($row)
        }
    }
    catch {}
    Write-Log -Message "enumerated $($services.Count) services" -LogFile $LogFile
    , $svcDetails | Export-CliXml -Path ($filename)
}
