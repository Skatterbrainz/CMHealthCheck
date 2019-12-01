function Test-Powershell64bit {
    Write-Output ([IntPtr]::size -eq 8)
}