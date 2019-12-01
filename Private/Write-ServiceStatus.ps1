function Write-ServiceStatus {
    param (
        [parameter(Mandatory)][string] $FileName,
        [parameter(Mandatory)][string] $TableName,
        [parameter()][string] $SiteCode,
        [parameter()][int] $NumberOfDays,
        [parameter()] $LogFile,
        [parameter()][string] $ServerName,
        [parameter()] $ContinueOnError = $true
    )
    Write-Log -Message "[function: write-servicestatus]" -LogFile $logfile

    $SiteInformation = Get-CmWmiObject -query "select Type from SMS_Site where ServerName = '$Server'" -namespace "Root\SMS\Site_$SiteCodeNamespace" -computerName $smsprovider -logfile $logfile
    if ($null -ne $SiteInformation) { $SiteType = $SiteInformation.Type }

    $WMISMSListRoles = Get-CmWmiObject -query "select distinct RoleName from SMS_SCI_SysResUse where NetworkOSPath = '\\\\$Server'" -computerName $smsprovider -namespace "root\sms\site_$SiteCodeNamespace" -logfile $logfile
    $SMSListRoles = @()
    foreach ($WMIServer in $WMISMSListRoles) { $SMSListRoles += $WMIServer.RoleName }
    Write-Log -Message "Roles discovered: " + $SMSListRoles -join(", ") -LogFile $logfile
    $Fields = @("ServiceName", "Status")
    $ServicesTable = New-CmDataTable -TableName $tableName -Fields $Fields

    if ($SMSListRoles -contains 'AI Update Service Point') {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "AI_UPDATE_SERVICE_POINT"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    if (($SMSListRoles -contains 'SMS Application Web Service') -or ($SMSListRoles -contains 'SMS Distribution Point') -or ($SMSListRoles -contains 'SMS Fallback Status Point') -or ($SMSListRoles -contains 'SMS Management Point') -or ($SMSListRoles -contains 'SMS Portal Web Site')  ) {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "IISADMIN"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "W3SVC"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS Component Server') {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "SMS_Executive"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS Site Server') {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "SMS_NOTIFICATION_SERVER"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "SMS_SITE_COMPONENT_MANAGER"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
        if ($SiteType -ne 1) {
            $row = $ServicesTable.NewRow()
            $row.ServiceName = "SMS_SITE_VSS_WRITER"
            $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
            $ServicesTable.Rows.Add($row)
        }
    }
    if ($SMSListRoles -contains 'SMS Software Update Point') {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "WsusService"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS SQL Server') {
        $row = $ServicesTable.NewRow()
        if ($SiteType -ne 1) {		
            $row.ServiceName = "$SQLServiceName"
        }
        else {
            $row.ServiceName = 'MSSQL$CONFIGMGRSEC'
        }
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "SQLWriter"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    if ($SMSListRoles -contains 'SMS SRS Reporting Point') {
        $row = $ServicesTable.NewRow()
        $row.ServiceName = "ReportServer"
        $row.Status = (Get-ServiceStatus -LogFile $logfile -ServerName $servername -ServiceName $row.ServiceName)
        $ServicesTable.Rows.Add($row)
    }
    , $ServicesTable | Export-CliXml -Path ($filename)
}