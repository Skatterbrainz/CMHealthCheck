function Invoke-Error {
    param (
        [parameter(Mandatory=$False, HelpMessage="Error or Exception Message")]
        [string] $Message = ""
    )
    if ([string]::IsNullOrEmpty($Message)) {
        Write-Log -Message $_.Exception.Message -Severity 3 -LogFile $logfile -ShowMsg
    }
    else {
        Write-Log -Message $Message -Severity 3 -LogFile $logfile -ShowMsg
    }
    Stop-Transcript -ErrorAction SilentlyContinue
}