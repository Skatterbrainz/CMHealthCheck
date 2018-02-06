function Set-ReplaceString {
    param (
		[parameter(Mandatory=$False)]
			[string] $Value,
		[parameter(Mandatory=$True)]
			[string] $SiteCode,
		[parameter(Mandatory=$False)]
			[int] $NumberOfDays = "",
		[parameter(Mandatory=$False)]
			[string] $ServerName = "",
		[parameter(Mandatory=$False)]
			[bool] $Space = $true
	)
	$return = $value
    $date = Get-Date
	if ($space) {	
		$return = $return -replace "\r\n", " " 
		$return = $return -replace "\r", " " 
		$return = $return -replace "\n", " " 
		$return = $return -replace "\s", " " 
		$return = $return -replace "\s{2}\b"," "
	}
	$return = $return -replace "@@SITECODE@@",$SiteCode
	$return = $return -replace "@@STARTMONTH@@",$date.ToString("01/MM/yyyy")
	$return = $return -replace "@@TODAYMORNING@@",$date.ToString("yyyy/MM/dd")
	$return = $return -replace "@@NUMBEROFDAYS@@",$NumberOfDays
	$return = $return -replace "@@SERVERNAME@@",$ServerName
	if ($space) {
		while (($return.IndexOf("  ") -ge 0)) { $return = $return -replace "  ", " " }
	}
	Write-Output $return
}