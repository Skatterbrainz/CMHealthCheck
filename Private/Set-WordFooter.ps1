function Set-WordFooter {
    Write-Log -Message "writing document footer content..."
    if ($Template -eq "") {
        $selection.HeaderFooter.Range.Text= "Copyright $([char]0x00A9) $((Get-Date).Year) - $CopyrightName"
    }
    $selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null
}
