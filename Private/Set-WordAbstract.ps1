function Set-WordAbstract {
	$absText1 = "This document provides a point-in-time report of the current state of the "
	$absText1 += "System Center Configuration Manager site environment for $CustomerName. "
	$absText1 += "For questions, concerns or comments, please consult the author of this "
	$absText1 += "assessment report."
	$absText2 = "This report was generated using CMHealthCheck $ModuleVer on $(Get-Date)."

	Write-WordText -WordSelection $selection -Text "Abstract" -Style "Heading 1" -NewLine $true
	Write-WordText -WordSelection $selection -Text $absText1 -NewLine $true
	Write-WordText -WordSelection $selection -Text $absText2 -NewLine $true
}