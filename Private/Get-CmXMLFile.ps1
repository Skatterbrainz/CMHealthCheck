function Get-CmXMLFile {
    param (
        [parameter(Mandatory=$True, HelpMessage="Path to file")]
            [ValidateNotNullOrEmpty()]
            [string] $Path,
        [parameter(Mandatory=$True, HelpMessage="File name")]
            [ValidateNotNullOrEmpty()]
            [string] $FileName
    )
    Write-Log -Message "function... Get-CmXMLFile ****" -LogFile $logfile
    $xfile = Join-Path -Path $Path -ChildPath $FileName
    try {
        $result = Import-CliXml -Path $xfile
    }
    catch {
        Write-Log -Message "File $xfile not found. No further action taken" -LogFile $logfile -Severity 3 -ShowMsg
        break
    }
    <#
    if (Test-Folder -Path $Path -Create $false) {
        $configfile = Join-Path -Path $Path -ChildPath "config.xml"
        if (!(Test-Path -Path $configfile)) {
            Write-Log -Message "File $configfile does not exist, no futher action taken" -Severity 3 -LogFile $logfile
            Stop-Transcript -ErrorAction SilentlyContinue
            break
        }
        else { 
            Write-Verbose "reading $configfile data"
            $ConfigTable = Import-CliXml -Path $configfile
        }
        
        if ($poshversion -ne 3) { $NumberOfDays = $ConfigTable.Rows[0].NumberOfDays }
        else { $NumberOfDays = $ConfigTable.NumberOfDays }
        $repfile = Join-Path -Path $Path -ChildPath "report.xml"
        if (!(Test-Path -Path $repfile)) {
            Write-Log -Message "File $repfile does not exist, no futher action taken" -Severity 3 -LogFile $logfile
            Stop-Transcript -ErrorAction SilentlyContinue
            break
        }
        else {
            $ReportTable = New-Object System.Data.DataTable 'ReportTable'
            $ReportTable = Import-CliXml -Path $repfile
        }
    }
    else {
        Write-Warning "Folder: $Path does not exist, no futher action taken"
        Stop-Transcript -ErrorAction SilentlyContinue
        break
    }
    #>
    Write-Output $result
}
    