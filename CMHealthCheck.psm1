$ScriptVersion = '1.0.3'
Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Private'),(Join-Path -Path $PSScriptRoot -ChildPath 'Public') -Filter '*.ps1' |
    ForEach-Object { . $_.FullName }
