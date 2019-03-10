<#
.DESCRIPTION
Write some text in Word
.NOTES
Fixed default from "No Spacing" to "Normal"
#>
function Write-WordText {
    param (
		$WordSelection,
        [parameter(Mandatory=$False)]
            [string] $Text = "",
        [parameter(Mandatory=$False)]
            [string] $Style = "Normal",
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