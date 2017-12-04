function Set-WordOptions {
    Write-Log -Message "configuring word options for current session" -LogFile $logfile
    $Word.Options.CheckGrammarAsYouType  = $False
    $Word.Options.CheckSpellingAsYouType = $False
    $Doc.Styles("Normal").Font.Size = $NormalFontSize
}