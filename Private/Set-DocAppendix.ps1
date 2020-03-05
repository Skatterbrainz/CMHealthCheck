function Set-DocAppendix {
	param ()
	$appendix = @(
		("ConfigMgr Hardware Recommendations", "https://technet.microsoft.com/en-us/library/mt589500.aspx#bkmk_ScaleSieSystems"),
		("ConfigMgr Supported Operating Systems", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/supported-operating-systems-for-site-system-servers"),
		("ConfigMgr Supported SQL Server Versions", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/support-for-sql-server-versions"),
		("ConfigMgr Internet-Based Client Management", "https://docs.microsoft.com/en-us/sccm/core/clients/manage/plan-internet-based-client-management"),
		("ConfigMgr Site Size and Scale Information", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/size-and-scale-numbers"),
		("ConfigMgr Support Lifecycle Information", "https://support.microsoft.com/en-us/lifecycle/search?alpha=Microsoft%20System%20Center%202012%20Configuration%20Manager"),
		("ConfigMgr Client Installation Properties", "https://docs.microsoft.com/en-us/sccm/core/clients/deploy/about-client-installation-properties"),
		("Best Practices for Managing Software Updates", "https://docs.microsoft.com/en-us/sccm/sum/plan-design/software-updates-best-practices"),
		("Deploy Windows 10 with MDT", "https://docs.microsoft.com/en-us/windows/deployment/deploy-windows-mdt/deploy-windows-10-with-the-microsoft-deployment-toolkit"),
		("Blogs - WindowsNoob", "https://www.windows-noob.com/forums/portal"),
		("Blogs - Deployment Research", "https://deploymentresearch.com/"),
		("Blogs - SC ConfigMgr", "https://www.scconfigmgr.com")
	)
	
	Write-Log -Message "inserting document Appendix..." -LogFile $logfile
	Write-WordText -WordSelection $selection -Text "Appendix A - Resource References" -Style "Heading 1" -NewLine $true
	$selection.TypeParagraph()

	foreach ($app in $appendix) {
		$caption = $app[0]
		$link = $app[1]
		Write-WordText -WordSelection $selection -Text $caption -NewLine $true
		Write-WordText -WordSelection $selection -Text $link -NewLine $true
		$selection.TypeParagraph()
	}
}