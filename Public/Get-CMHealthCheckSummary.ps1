function Get-CMHealthCheckSummary {
    [CmdletBinding()]
    param (
        [parameter(HelpMessage="Reporting scope")]
			[ValidateSet('All','AD','CM','SQL')]
			[string] $ReportScope = 'All',
        [parameter(HelpMessage="Location to save report file")]
			[string] $OutputFolder = $(Join-Path -Path $env:USERPROFILE -ChildPath "Documents"),
		[parameter()][switch] $Detailed
    )
	$divtag = "table1"
    try {
        $reportFile  = Join-Path $env:USERPROFILE -ChildPath "Documents\cmhealthcheck_summary.htm"
		Write-Verbose "report file: $reportFile"
        $reportTitle = "CMHealthCheck Summary"
        $content = ""
        if ($ReportScope -in @('All','AD')) {
            $forestInfo  = Get-CmhAdForestInfo
            $adsites     = Get-CmhAdSites
            $adsitelinks = Get-CmhAdSiteLinks
            $adusers     = Get-CmhAdUsers
            $adgroups    = Get-CmhAdGroups
            $adcomps     = Get-CmhAdComputers
            $smcont      = Test-CmhAdContainer
			$smextended  = Test-CmhAdSchemaExtension
			$content += "<h1>Active Directory</h1>"
            $content += "<h2>Forest and Domain</h2>"
            $content += $forestinfo | ConvertTo-HtmlPivot -DivID $divtag

            $content += "<h2>Accounts</h2>
<table id=table1>
<tr><th>ObjectCategory</th><th>Count</th></tr>
<tr><td>Users</td><td>$($adusers.Count)</td></tr>
<tr><td>Groups</td><td>$($adgroups.Count)</td></tr>
<tr><td>Computers</td><td>$($adcomps.Count)</td></tr>
</table>"

            $content += "<h2>Sites</h2>"
            $content += $adsites | ConvertTo-HtmlPivot -DivID $divtag

            $content += "<h2>Site Links</h2>"
            $content += $adsitelinks | ConvertTo-HtmlPivot -DivID $divtag

            $content += "<h2>Configuration Manager AD Schema</h2>
<table id=$divtag>
<tr><th>Condition</th><th>Status</th></tr>
<tr><td>AD Schema extension applied</td><td>$smextended</td></tr>
<tr><td>System Management Container created</td><td>$smcont</td></tr>
</table>"
        }
        if ($ReportScope -in @('All','CM')) {
			Write-Verbose "collecting configmgr data..."
			$content += "<h1>Configuration Manager</h1>"
			$mp = Get-CmhCmClientMP
			Write-Verbose "mp: $mp"
			$sitecode = Get-CmhCmSiteCode
			Write-Verbose "sitecode: $sitecode"
            $content += "<h2>Configuration Manager Client</h2>
<table id=$divtag>
<tr><th>Condition</th><th>Status</th></tr>
<tr><td>Site Code</td><td>$sitecode</td></tr>
<tr><td>Management Point</td><td>$mp</td></tr>
</table>"

			$content += "<h2>Site Systems and Roles</h2>"
			$content += (Get-CmhCmSiteSystemRoles | ConvertTo-Html -Fragment) -replace "<table>","<table id=$divtag>"

        }
        if ($ReportScope -in @('All','SQL')) {
			$content += "<h1>SQL Server</h1>"
			$content += "<h2>ConfigMgr SQL Server Host</h2>"
			$mphost   = Get-CmhCmClientMP
			$sitecode = Get-CmhCmSiteCode
			$dbhost   = Get-CmhCmSiteSystemRoles | Where-Object {$_.RoleName -eq 'SMS SQL Server'} | Select-Object -ExpandProperty Name
			Write-Verbose "sql host: $dbhost"
			$content += Get-DbaComputerSystem -ComputerName $dbhost | ConvertTo-HtmlPivot -DivID $divtag

			$content += "<h2>Database: CM_$sitecode</h2>"
			$content += Get-DbaComputerSystem -ComputerName $dbhost | ConvertTo-HtmlPivot -DivID $divtag

			$content += "<h2>Database: CM_$sitecode Status</h2>"
			$content += Get-DbaDbState -SqlInstance $dbhost -Database "CM_$sitecode" | ConvertTo-HtmlPivot -DivID $divtag

			if ($Detailed) {
				$content += "<h2>Database Files: CM_$sitecode</h2>"
				$content += Get-DbaDbFile -SqlInstance $dbhost -Database "CM_$sitecode" | ForEach-Object {ConvertTo-HtmlPivot -InputObject $_ -DivID $divtag}
				$content += "<h2>SQL Server DBCC Status</h2>"
				$content += (Get-DbaDbccMemoryStatus -SqlInstance $dbhost | ConvertTo-Html -Fragment) -replace "<table>","<table id=$divtag>"
			}
        }
		$content += "<p class=footer>Generated: $(Get-Date) `&copy`; 2019 skatterbrainz. Refer to LICENSE for usage rights and conditions.</p>"
        Write-CmhHtml -Title $reportTitle -BodyContent $content -FilePath $reportFile
		Write-Host "Report saved to $reportFile" -ForegroundColor Cyan
    }
    catch {
		Write-Warning "Error: $($Error[0].Exception.Message)"
	}
}
