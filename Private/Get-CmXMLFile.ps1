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
    Write-Output $result
}
    