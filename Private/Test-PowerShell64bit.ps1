function Test-Powershell64bit {
    <#
    .SYNOPSIS
    Check if PowerShell is running in 64-bit context
    
    .DESCRIPTION
    Return $True if PowerShell is running in 64-bit context
    
    .EXAMPLE
    if (!(Test-PowerShell64bit)) {
        Write-Error "You are screwed"
        break
    }
    
    .NOTES
    General notes
    #>
    Write-Output ([IntPtr]::size -eq 8)
}