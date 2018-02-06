function Write-WordText {
    param (
		$WordSelection,
        [parameter(Mandatory=$False)]
            [string] $Text = "",
        [parameter(Mandatory=$False)]
            [string] $Style = "No Spacing",
		$Bold    = $false,
		$NewLine = $false,
		$NewPage = $false
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