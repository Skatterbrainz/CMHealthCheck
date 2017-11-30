function Get-CmSiteInstallPath {
    <#
    .SYNOPSIS
        Get-CmSiteInstallPath returns [string] path to the base installation
        of System Center Configuration Manager on the site server.
    .DESCRIPTION
        Returns the full SCCM installation path using a registry query.
    #>
	Write-Log -Message "getting configmgr installation path" -Logfile $logfile
	try {
		$x = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\setup"
		Write-Output $x.'Installation Directory'
	}
	catch {}
}
