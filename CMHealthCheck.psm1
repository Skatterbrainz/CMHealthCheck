$Script:ScriptVersion   = '1.0.1'
$(Get-ChildItem "$PSScriptRoot" -Recurse -Include "*.ps1").foreach{. $_.FullName}