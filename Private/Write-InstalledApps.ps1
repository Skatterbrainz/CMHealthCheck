function Write-InstalledApps {
    param (
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [string] $Filename,
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [string] $TableName,
        [parameter(Mandatory=$True)]
            [string] $SiteCode,
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [string] $ServerName,
        [parameter(Mandatory=$False)]
            [string] $LogFile,
        [parameter(Mandatory=$False)]
            [bool] $ContinueOnError
    )
    Write-Log -Message "function... Write-InstalledApps ****" -LogFile $logfile
    Write-Log -Message "filename... $filename" -LogFile $LogFile
    Write-Log -Message "server..... $ServerName" -LogFile $LogFile
    try {
        $Apps = Get-WmiObject -Class "Win32_Product" -ComputerName $ServerName -ErrorAction Stop
    }
    catch {
        if ($ContinueOnError -eq $True) {
            Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate software' -LogFile $LogFile
        }
        else {
            Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate software' -Severity 3 -LogFile $LogFile
            return
        }
    }
    if ($Apps -eq $null) {
        Write-Log -Message "found NO installed applications (aborting)" -LogFile $LogFile
        return
    }
    Write-Log -Message "found $($Apps.Count) installed applications" -LogFile $LogFile
    $Fields=@("Name","Version","Vendor")
    $AppDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    foreach ($app in $Apps) {
        $appname = $app.Name
        $appver  = $app.Version 
        $appven  = $app.Vendor 
        $row = $AppDetails.NewRow()
        $row.Name = $appname
        $row.Version = $appver
        $row.Vendor = $appven
        $AppDetails.Rows.Add($row)
    }
    , $AppDetails | Export-CliXml -Path ($filename)
}