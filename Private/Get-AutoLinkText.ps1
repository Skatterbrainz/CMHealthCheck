function Get-AutoLinkText {
	param (
		[parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[string] $StringValue,
		[parameter()][switch] $NewPage
	)
	Write-Log -Message "(Get-AutoLinkText)" -LogFile $logfile
	$temp = @()
	$tokens = $StringValue.split(' ')
	$tokens | ForEach-Object {
		if ($_.StartsWith('http')) {
			if ($NewPage) {
				$temp += "<a href=`"$_`" target=`"blank`">$_</a>"
			} else {
				$temp += "<a href=`"$_`">$_</a>"
			}
		} elseif ($_.StartsWith('(http') -and $_.EndsWith(')')) {
			if ($NewPage) {
				$temp += "(<a href=`"$_`" target=`"blank`">$_</a>)"
			} else {
				$temp += "(<a href=`"$_`">$_</a>)"
			}
		} elseif ($_.StartsWith('(http') -and $_.EndsWith(').')) {
			$tx = ($_.Replace('(','')).Replace(').','')
			if ($NewPage) {
				$temp += "(<a href=`"$tx`" target=`"blank`">$tx</a>)."
			} else {
				$temp += "(<a href=`"$tx`">$tx</a>)."
			}
		} else {
			$temp += $_
		}
	}
	$temp -join ' '
}