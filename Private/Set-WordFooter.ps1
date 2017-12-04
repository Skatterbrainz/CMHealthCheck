function Set-WordFooter {
    $selection.HeaderFooter.Range.Text= "Copyright $([char]0x00A9) $((Get-Date).Year) - $CopyrightName"
    $selection.HeaderFooter.PageNumbers.Add(2) | Out-Null
}
