function Write-WordText {
    param (
		[parameter()] $WordSelection,
        [parameter()][string] $Text = "",
        [parameter()][string] $Style = "Normal",
		[parameter()] $Bold    = $false,
		[parameter()] $NewLine = $false,
		[parameter()] $NewPage = $false
	)
	$texttowrite = ""
	$wordselection.Style = $Style
    if ($bold) { 
		$wordselection.Font.Bold = 1 
	}
	else { 
		$wordselection.Font.Bold = 0
	}
	$texttowrite += $text
	$wordselection.TypeText($text)
	If ($newline) { $wordselection.TypeParagraph() }
	If ($newpage) { $wordselection.InsertNewPage() }
}