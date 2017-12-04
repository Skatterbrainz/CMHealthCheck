Function Write-Log {
    [CmdletBinding()]
    param (
		[parameter(Mandatory=$False)]
			[ValidateNotNullOrEmpty()]
			[String] $Message = "",
        [parameter(Mandatory=$False)][int] $Severity = 1,
        [parameter(Mandatory=$False)][string] $LogFile = '',
        [parameter(Mandatory=$False)][switch] $ShowMsg
        
    )
    switch ($Severity) {
        1 {$Category='Info'; break}
        2 {$Category='Warning'; break}
        3 {$Category='Error'; break}
    }
    $MsgTxt = "$(Get-Date -f 'yyyy-M-dd HH:mm:ss')  $Category  $Message"
    if (($logfile -ne $null) -and ($LogFile -ne '')) {
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
