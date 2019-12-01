function Get-TimeOffset {
    param (
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][datetime] $StartTime
    )
    $secs = ((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds).ToString()
    $ts   = [timespan]::FromSeconds($secs)
    Write-Output $ts.ToString("hh\:mm\:ss")
}