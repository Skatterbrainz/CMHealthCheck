function Write-LocalGroups {
    param (
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $Filename,
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $TableName,
        [parameter(Mandatory)][string] $SiteCode,
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $ServerName,
        [parameter()][string] $LogFile,
        [parameter()][bool] $ContinueOnError
    )
    Write-Log -Message "function... Write-LocalGroups ****" -LogFile $logfile
    $ServerShortName = ($ServerName -split '\.')[0]
    try {
        $GroupsList = Get-CimInstance -ClassName "Win32_Group" -ComputerName $ServerName -Filter "Domain='$ServerShortName'" -ErrorAction Stop
    }
    catch {
        Write-Log -Category 'Error' -Message 'cannot connect to $ServerName to enumerate local security groups'
        return
    }
    if ($null -eq $GroupsList) { return }
    $Fields = @("Name","Description","Members")
    $GroupDetails = New-CmDataTable -TableName $tableName -Fields $Fields
    foreach ($group in $GroupsList) {
        $gn = $group.Name
        $gd = $group.Description
        Write-Log -Message "group name... $gn" -LogFile $LogFile
        # get netbios (short) server name
        $SSname = ($ServerName.Split('\.'))[0]
        Write-Log -Message "getting members..." -LogFile $LogFile
        $wmiquery = "select * from Win32_GroupUser where GroupComponent=`"Win32_Group.Domain='$SSName',Name='$gn'`""
        Write-Log -Message "wmi query.... $wmiquery" -LogFile $LogFile
        try {
            $acct = Get-CimInstance -ComputerName $ServerName -Query $wmiquery -ErrorAction Stop
            $arr = @()
            if ($null -ne $acct) {
                foreach ($item in $acct) {
                    $data   = $item.PartComponent -split "\,"
                    $domain = ($data[0] -split "=")[1]
                    $name   = ($data[1] -split "=")[1]
                    $arr   += ("$domain\$name").Replace("""","")
                    [Array]::Sort($arr)
                }
                if ($arr.Count -gt 0) {
                    Write-Log -Message "member count... $($arr.Count)" -LogFile $LogFile
                    [string]$members = ($arr -join ", ")
                }
                else {
                    $members = '(no members)'
                }
            }
        }
        catch {
            if ($ContinueOnError -eq $True) {
                Write-Log -Category 'Error' -Message $_.Exception.Message -Severity 2 -LogFile $logfile
            }
            else {
                Write-Log -Category 'Error' -Message "Terminating Error: $($_.Exception.Message)" -Severity 3 -LogFile $logfile
                return
            }          
            $members = '(no members - failed to enumerate)'
        }
        Write-Log -Message "members...... $members" -LogFile $LogFile
        $row             = $GroupDetails.NewRow()
        $row.Name        = $gn
        $row.Description = $gd
        $row.Members     = $members
        $GroupDetails.Rows.Add($row)
    } # foreach group
    Write-Log -Message "enumerated $($GroupsList.Count) groups" -LogFile $LogFile
    , $GroupDetails | Export-CliXml -Path ($filename)
}