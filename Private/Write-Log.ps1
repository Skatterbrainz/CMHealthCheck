Function Write-Log {
    [CmdletBinding()]
    param (
		[parameter()][ValidateNotNullOrEmpty()][String] $Message = "",
        [parameter()][int] $Severity = 1,
        [parameter()][string] $LogFile = '',
        [parameter()][switch] $ShowMsg
        
    )
    switch ($Severity) {
        1 { $Category='Info' }
        2 { $Category='Warning' }
        3 { $Category='Error' }
    }
    $MsgTxt = "$(Get-Date -f 'yyyy-M-dd HH:mm:ss')  $Category  $Message"
    if (![string]::IsNullOrEmpty($logfile)) {
		$MsgTxt | Out-File -FilePath $LogFile -Append -NoClobber -Encoding Default
    }
    if ($showmsg) {
        switch ($Severity) {
            3 { Write-Host $Message -ForegroundColor Red }
            2 { Write-Host $Message -ForegroundColor Yellow }
            1 { Write-Host $Message }
        }
    }
	else {
		Write-Verbose $MsgTxt
	}
}
