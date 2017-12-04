function Get-WordTempSource {
    <#
    .SYNOPSIS
    Copy Source Document File to Destination
    
    .DESCRIPTION
    Copies a source DOCX file to a temporary name and returns the new filename
    
    .PARAMETER SourceFile
    Path and name of source document file
    
    .EXAMPLE
    $newfile = Get-WordTempSource -SourceFile "c:\files\myfile.docx"
    $newfile == "c:\users\johndoe\documents\cmhealthreport.docx"
    
    .NOTES
    #>
    param (
        [parameter(Mandatory=$True, HelpMessage="Name of Template File")]
        [ValidateNotNullOrEmpty()]
        [string] $SourceFile
    )
    if (Test-Path -Path $SourceFile) {
        $newFile = Join-Path -Path $OutputFolder -ChildPath $TempFilename
        Write-Log -Message "copying source [$Template] to temp file [$newFile]..." -LogFile $logfile
        try {
            Copy-Item -Path $Template -Destination $newFile -ErrorAction Stop
            $result = $True
        }
        catch {
            Write-Log -Message "ERROR: Failed to clone template from $Template" -Severity 3 -LogFile $logfile
            break
        }
    }
    Write-Output $newFile
}