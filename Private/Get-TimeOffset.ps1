function Get-TimeOffset {
    <#
    .SYNOPSIS
    Return Elapsed Time
    
    .DESCRIPTION
    Return Elapsed Time from a given starting time
    
    .PARAMETER StartTime
    DateTime value (e.g. Get-Date)
    
    .EXAMPLE
    $t1 = (Get-Date)
    # ...
    Write-Host Get-TimeOffset -StartTime $t1
    
    .NOTES
    General notes
    #>

    param (
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [datetime] $StartTime
    )
    $secs = ((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds).ToString()
    $ts   = [timespan]::FromSeconds($secs)
    Write-Output $ts.ToString("hh\:mm\:ss")
}