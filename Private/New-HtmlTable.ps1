function New-HtmlTableBegin {
	[CmdletBinding()]
	param (
		[parameter()][string] $Caption = "",
		[parameter()][string] $CaptionStyle = "h2", 
		[parameter()][string] $TableClass = "reportTable",
		[parameter()][string] $HeadingStyle = "",
		[parameter()][string] $HeadingNames
	)
	Write-Log -Message "--- function: New-HtmlTableBegin" -LogFile $logfile
	Write-Log -Message "--- tableclass`:$TableClass headingstyle`:$HeadingStyle" -LogFile $logfile

	if (!([string]::IsNullOrEmpty($Caption))) {
		$Caption = "`n<$CaptionStyle>$Caption</$CaptionStyle>"
	}
	$result = "`n$Caption <table class=`"$TableClass`">"

	$result += "<tr>"
	foreach ($item in $HeadingNames.Split(',')) {
		if ($item -match '=') {
			$text = $item.Split('=')[0]
			$cwid = $item.Split('=')[1]
			$result += "<th class=`"columnstyle1`" style=`"width`:$($cwid)px`">$text</th>"
		}
		else {
			$result += "<th class=`"columnstyle1`">$item</th>"
		}
	}
	$result += "</tr>"
	Write-Output $result
}

function New-HtmlTableEnd {
	param (
		[string] $TableData
	)
	Write-Log -Message "--- function: New-HtmlTableEnd" -LogFile $logfile
	Write-Output $TableData += "`n</table>"
}

function New-HtmlTableBlock {
	param (
		[parameter()][string] $Caption = "",
		[parameter()][string] $CaptionStyle = "h2",
		[parameter()][string] $TableClass = "reportTable",
		[parameter()][string] $HeadingStyle = "",
		[parameter()][string] $HeadingNames,
		[parameter()][int] $Rows = 1
	)
	Write-Log -Message "--- function: New-HtmlTableBlock" -LogFile $logfile
	$result = New-HtmlTableBegin -Caption $Caption -CaptionStyle $CaptionStyle -TableClass $TableClass -HeadingStyle $HeadingStyle -HeadingNames $HeadingNames
	$columns = $HeadingNames.Split(',').Count
	for ($row=1; $row -le $rows; $row++) {
		# alternate row styles (even/odd)
		if ($row % 2 -eq 0) {
			$result += "<tr class=`"rowstyle1`">"
		}
		else {
			$result += "<tr class=`"rowstyle2`">"
		}
		for ($col=1;$col -le $columns;$col++) {
			$result += "<td>&nbsp;</td>"
		}
		$result += "</tr>"
	}
	$result += "`n</table>"
	Write-Output $result
}

function New-HtmlTableVertical {
	param (
		[parameter()][string] $Caption = "",
		[parameter()][string] $CaptionStyle = "h2",
		[parameter()][string] $TableClass = "reportTable",
		[parameter()][hashtable] $TableHash
	)
	Write-Log -Message "--- function: New-HtmlTableVertical" -LogFile $logfile
	if (!([string]::IsNullOrEmpty($Caption))) {
		$Caption = "`n<$CaptionStyle>$Caption</$CaptionStyle>"
	}
	$result = "`n$Caption <table class=`"$TableClass`">"
	foreach ($key in $TableHash.Keys) {
		$val  = $TableHash.Item($key)
		$col1 = "<td class=`"columnstyle1`">$key</td>"
		$col2 = "<td class=`"columnstyle2`">$val</td>"
		$result += "<tr>$($col1)$($col2)</tr>"
	}
	$result += "`n</table>"
	Write-Output $result
}
