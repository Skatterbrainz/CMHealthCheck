function Set-DocAppendix {
    Write-WordText -WordSelection $selection -Text "Appendix A - Resource References" -Style "Heading 1" -NewLine $true
#    Write-WordText -WordSelection $selection -Text "(insert links to relevant documentation, resources, etc.)" -NewLine $true
    Write-WordText -WordSelection $selection -Text "ConfigMgr Hardware Recommendations" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://technet.microsoft.com/en-us/library/mt589500.aspx#bkmk_ScaleSieSystems" -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Supported Operating Systems" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/supported-operating-systems-for-site-system-servers " -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Supported SQL Server Versions" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/support-for-sql-server-versions " -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Internet-Based Client Management" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/core/clients/manage/plan-internet-based-client-management " -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Site Size and Scale Information" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/size-and-scale-numbers " -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Support Lifecycle Information" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://support.microsoft.com/en-us/lifecycle/search?alpha=Microsoft%20System%20Center%202012%20Configuration%20Manager " -NewLine $True

    Write-WordText -WordSelection $selection -Text "ConfigMgr Client Installation Properties" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/core/clients/deploy/about-client-installation-properties " -NewLine $True

    Write-WordText -WordSelection $selection -Text "Best Practices for Managing Software Updates" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://docs.microsoft.com/en-us/sccm/sum/plan-design/software-updates-best-practices " -NewLine $True

    Write-WordText -WordSelection $selection -Text "Blogs - WindowsNoob" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://www.windows-noob.com/forums/portal/ " -NewLine $True

    Write-WordText -WordSelection $selection -Text "Blogs - SC ConfigMgr" -NewLine $True
    Write-WordText -WordSelection $selection -Text "https://www.scconfigmgr.com " -NewLine $True
}