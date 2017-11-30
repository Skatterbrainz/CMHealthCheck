Function Write-Log {
    [CmdletBinding()]
    param (
		[parameter(Mandatory=$True)]
			[ValidateNotNullOrEmpty()]
			[String] $Message,
        [parameter(Mandatory=$False)][int] $Severity = 1,
        [parameter(Mandatory=$False)][string] $LogFile = '',
        [parameter(Mandatory=$False)][switch] $ShowMsg
        
    )
    #$TimeZoneBias = Get-WmiObject -Query "Select Bias from Win32_TimeZone"
    #$Date  = Get-Date -Format "HH:mm:ss.fff"
    #$Date2 = Get-Date -Format "MM-dd-yyyy"
    if (($logfile -ne $null) -and ($logfile -ne '')) {
		#"<![LOG[$Message]LOG]!><time=`"$date+$($TimeZoneBias.Bias)`" date=`"$date2`" component=`"$component`" context=`"`" type=`"$severity`" thread=`"`" file=`"`">" | 
		#	Out-File -FilePath $logfile -Append -NoClobber -Encoding Default
		switch ($Severity) {
			1 {$Category='Info'; break}
			2 {$Category='Warning'; break}
			3 {$Category='Error'; break}
		}
		$Msg = "$(Get-Date -f 'yyyy-M-dd HH:mm:ss')  $Category  $Message"
		$Msg | Out-File -FilePath $LogFile -Append -NoClobber -Encoding Default
    }
    if ($showmsg) {
        switch ($Severity) {
            3 { Write-Host $Message -ForegroundColor Red }
            2 { Write-Host $Message -ForegroundColor Yellow }
            1 { Write-Host $Message }
        }
    }
	else {
		Write-Verbose $Msg
	}
}
