#requires -version 3
<#
.SYNOPSIS
    Export-CMHealthcheck.ps1 reads the output from Get-CM-Inventory.ps1 to generate a
    final report using Microsoft Word (2010, 2013, 2016)

.DESCRIPTION
    Export-CMHealthcheck.ps1 reads the output from Get-CM-Inventory.ps1 to generate a
    final report using Microsoft Word (2010, 2013, 2016)

.PARAMETER ReportFolder
    [string] [required] Path to output data folder

.PARAMETER Detailed
    [switch] [optional] Collect more granular data for final reporting

.PARAMETER Healthcheckfilename
    [string] [optional] healthcheck configuration file name
	default   = "https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml"
	alternate = ".\cmhealthcheck.xml"

.PARAMETER Healthcheckdebug
    [boolean] [optional] Enable verbose output (or use -Verbose)

.PARAMETER CoverPage
    [string] [optional] 
    default = "Slice (Light)"

.PARAMETER CustomerName
    [string] [optional] Name of customer
    default = "Company"

.PARAMETER AuthorName
    [string] [optional] Name of report author
    default = "Author"

.PARAMETER Overwrite
    [switch] [optional] Overwrite existing report file if found

.NOTES
    Thanks to:
    Base script (the hardest part) created by Rafael Perez (www.rflsystems.co.uk)
    Word functions copied from Carl Webster (www.carlwebster.com)
    Word functions copied from David O'Brien (www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/)

.EXAMPLE
    Option 1: powershell.exe -ExecutionPolicy Bypass .\Export-CMHealthcheck [Parameters]
    Option 2: Open Powershell and execute .\Export-CMHealthcheck [Parameters]
    Option 3: .\Export-CMHealthCheck -ReportFolder "2017-05-17\cm1.contoso.com" -Detailed -CustomerName "ACME" -AuthorName "David Stein" -Overwrite -Verbose

#>

function Export-CMHealthCheck {
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $True, HelpMessage = "Collected data folder")] 
            [ValidateNotNullOrEmpty()]
            [string] $ReportFolder,
        [Parameter (Mandatory = $False, HelpMessage = "Export full data, not only summary")] 
            [switch] $Detailed,
        [parameter (Mandatory = $False, HelpMessage = "Word Template cover page name")] 
            [string] $CoverPage = "Slice (Light)",
        [parameter (Mandatory = $False, HelpMessage = "Customer company name")] 
            [string] $CustomerName = "Customer Name",
        [parameter (Mandatory = $False, HelpMessage = "Author's full name")] 
            [string] $AuthorName = "Your Name",
        [parameter (Mandatory = $False, HelpMessage = "Footer text")]
            [string] $CopyrightName  = "Your Company Name",
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck query file name")] 
            [string] $Healthcheckfilename = 'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml',
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck messages file name")]
            [string] $MessagesFilename = 'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/Messages.xml',
        [Parameter (Mandatory = $False, HelpMessage = "Debug more?")] 
            $Healthcheckdebug = $False,
        [parameter (Mandatory = $False, HelpMessage = "Overwrite existing report file")]
            [switch] $Overwrite
    )
    $time1 = Get-Date -Format "hh:mm:ss"
    Start-Transcript -Path ".\_logs\export-reportfile.log" -Append
    $ScriptVersion  = "1710.01"
    $bLogValidation = $False
    $bAutoProps     = $True
    $currentFolder  = $PWD.Path
    $NormalFontSize = 10
    $poshversion    = $PSVersionTable.PSVersion.Major
    $osversion      = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    $FormatEnumerationLimit = -1
    
    if ($healthcheckdebug -eq $true) { $PSDefaultParameterValues = @{"*:Verbose"=$True}; $currentFolder = "C:\Temp\CMHealthCheck\" }
    $logFolder = $currentFolder + "\_Logs\"
    if (-not (Test-Path $logFolder)) {
        Write-Error "$logFolder was not found!"
        break
    }
    if ($reportFolder.substring($reportFolder.length-1) -ne '\') { $reportFolder+= '\' }
    $component = ($MyInvocation.MyCommand.Name -replace '.ps1', '')
    $logfile = $logFolder + $component + ".log"
    $Error.Clear()

    Write-Log -Message "==========" -LogFile $logfile -ShowMsg $false
    Write-Log -Message "Script Version: $ScriptVersion" -LogFile $logfile
    Write-Log -Message "Windows Version: $osversion" -LogFile $logfile
    Write-Log -Message "Running Powershell version: $poshversion" -LogFile $logfile
    Write-Log -Message "Running Powershell 64 bits" -LogFile $logfile
    Write-Log -Message "Report Folder: $reportFolder" -LogFile $logfile
    Write-Log -Message "Detailed Report: $detailed" -LogFile $logfile
    Write-Verbose "Export-CMHealthCheck $ScriptVersion"
    Write-Verbose "Current Folder: $currentFolder"
    Write-Verbose "Log Folder: $logFolder"
    Write-Verbose "Log File: $logfile" 
    Write-Verbose "Healthcheck Data File: $Healthcheckfilename"

    Write-Output "script version: $ScriptVersion"
    $poshversion = $PSVersionTable.PSVersion.Major
    
    Write-Verbose "info: connecting to Microsoft Word..."
    try {
        $Word = New-Object -ComObject "Word.Application" -ErrorAction Stop
    }
    catch {
        Write-Warning "Error: This script requires Microsoft Word"
        break
    }

    if ($Healthcheckfilename.StartsWith('http')) {
        Write-Verbose "importing xml from remote URI: $healthcheckfilename"
        try {
            [xml]$HealthCheckXML = Invoke-RestMethod -Uri $Healthcheckfilename
        }
        catch {
            Write-Error "Failed to import data from Uri: $HealthcheckFilename"
            Write-Warning "If no Internet access is allowed, use -HealthcheckFilename '.\cmhealthcheck.xml'"
            break
        }
        Write-Verbose "configuration XML data loaded successfully"
    }
    else {
        Write-Verbose "importing Configuration xml from local file: $healthcheckfilename"
        if (!(Test-Path -Path $healthcheckfilename)) {
            Write-Warning "File $healthcheckfilename does not exist, no futher action taken"
            break
        }
        else { 
            try {
                [xml]$HealthCheckXML = Get-Content ($healthcheckfilename) 
            }
            catch {
                Write-Error "Failed to import data from local file: $HealthcheckFilename"
                break
            }
            Write-Verbose "configuration XML data loaded successfully"
        }
    }
    
    if ($MessagesFilename.StartsWith('http')) {
        Write-Verbose "importing Messages xml from remote URL: $MessagesFilename"
        try {
            [xml]$MessagesXML = Get-XmlUrlContent -Url $MessagesFilename
        }
        catch {
            Write-Error "Failed to import data from Uri: $MessagesFilename"
            Write-Warning "If no Internet access is allowed, use -MessagesFilename '.\messages.xml'"
            break
        }
        Write-Verbose "Messages XML data loaded successfully"
    }
    else {
        if (!(Test-Path -Path ".\Messages.xml")) {
            Write-Warning "File Messages.xml does not exist, no futher action taken"
            break
        }
        else { 
            Write-Verbose "reading messages.xml data"
            try {
                [xml]$MessagesXML = Get-Content '.\Messages.xml'
            }
            catch {
                Write-Error "Failed to import data from local file: $MessagesFilename"
                break
            }
        }
        Write-Verbose "Messages XML data loaded successfully"
    }
    
    if ($HealthCheckXML -and $MessagesXML) {
        if (Test-Folder -Path $logFolder) {
            try {
                New-Item ($logFolder + 'Test.log') -Type File -Force | Out-Null 
                Remove-Item ($logFolder + 'Test.log') -Force | Out-Null 
            }
            catch {
                Write-Warning "Unable to read/write file on $logFolder folder, no futher action taken"
                break    
            }
        }
        else {
            Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
            break
        }
        $bLogValidation = $true
    
        if (Test-Folder -Path $reportFolder -Create $false) {
            if (!(Test-Path -Path ($reportFolder + "config.xml"))) {
                Write-Log -Message "File $($reportFolder)config.xml does not exist, no futher action taken" -Severity 3 -LogFile $logfile
                break
            }
            else { 
                Write-Verbose "reading config.xml data"
                $ConfigTable = Import-CliXml -Path ($reportFolder + "config.xml") 
            }
            
            if ($poshversion -ne 3) { $NumberOfDays = $ConfigTable.Rows[0].NumberOfDays }
            else { $NumberOfDays = $ConfigTable.NumberOfDays }
            
            if (!(Test-Path -Path ($reportFolder + "report.xml"))) {
                Write-Log -Message "File $($reportFolder)report.xml does not exist, no futher action taken" -Severity 3 -LogFile $logfile
                break
            }
            else {
                 $ReportTable = New-Object System.Data.DataTable 'ReportTable'
                $ReportTable = Import-CliXml -Path ($reportFolder + "report.xml")
            }
        }
        else {
            Write-Warning "Folder: $reportFolder does not exist, no futher action taken"
            break
        }
        
        if (!(Test-Powershell64bit)) {
            Write-Log -Message "Powershell is not 64bit, no futher action taken" -Severity 3 -LogFile $logfile
            break
        }
        
        $wordVersion = $Word.Version
        Write-Log -Message "Word Version: $WordVersion" -LogFile $logfile	
        Write-Verbose "info: Microsoft Word version: $WordVersion"
        if ($WordVersion -ge "16.0") {
            $TableStyle = "Grid Table 4 - Accent 1"
            $TableSimpleStyle = "Grid Table 4 - Accent 1"
        }
        elseif ($WordVersion -eq "15.0") {
            $TableStyle = "Grid Table 4 - Accent 1"
            $TableSimpleStyle = "Grid Table 4 - Accent 1"
        }
        elseif ($WordVersion -eq "14.0") {
            $TableStyle = "Medium Shading 1 - Accent 1"
            $TableSimpleStyle = "Light Grid - Accent 1"
        }
        else { 
            Write-Log -Message "This script requires Word 2010 to 2016 version, no further action taken" -Severity 3 -LogFile $logfile 
            break
        }
    
        $Word.Visible = $True
        $Doc = $Word.Documents.Add()
        $Selection = $Word.Selection
        
        Write-Verbose "info: disabling real-time spelling and grammar check"
        $Word.Options.CheckGrammarAsYouType  = $False
        $Word.Options.CheckSpellingAsYouType = $False
        $Doc.Styles("Normal").Font.Size = $NormalFontSize
        
        Write-Verbose "info: loading default building blocks template"
        $word.Templates.LoadBuildingBlocks() | Out-Null	
        $BuildingBlocks = $word.Templates | Where-Object {$_.name -eq "Built-In Building Blocks.dotx"}
        $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
        
        if ($doc -eq $null) {
            Write-Error "Failed to obtain handle to Word document"
            break
        }
        if ($bAutoProps -eq $True) {
            Write-Verbose "info: setting document properties"
            $doc.BuiltInDocumentProperties("Title")    = "System Center Configuration Manager HealthCheck"
            $doc.BuiltInDocumentProperties("Subject")  = "Prepared for $CustomerName"
            $doc.BuiltInDocumentProperties("Author")   = $AuthorName
            $doc.BuiltInDocumentProperties("Company")  = $CopyrightName
            $doc.BuiltInDocumentProperties("Category") = "HEALTHCHECK"
            $doc.BuiltInDocumentProperties("Keywords") = "sccm,healthcheck,systemcenter,configmgr,$CustomerName"
        }
    
        Write-Verbose "info: inserting document parts"
        $part.Insert($selection.Range,$True) | Out-Null
        $selection.InsertNewPage()
        
        Write-Verbose "info: inserting table of contents"
        $toc = $BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
        $toc.Insert($selection.Range,$True) | Out-Null
    
        $selection.InsertNewPage()
    
        $currentview = $doc.ActiveWindow.ActivePane.view.SeekView
        $doc.ActiveWindow.ActivePane.view.SeekView = 4
        $selection.HeaderFooter.Range.Text= "Copyright $([char]0x00A9) $((Get-Date).Year) - $CopyrightName"
        $selection.HeaderFooter.PageNumbers.Add(2) | Out-Null
        $doc.ActiveWindow.ActivePane.view.SeekView = $currentview
        $selection.EndKey(6,0) | Out-Null
    
        $absText = "This document provides a point-in-time inventory and analysis of the "
        $absText += "System Center Configuration Manager site environment for $CustomerName. "
        $absText += "For questions, concerns or comments, please consult the $CopyrightName "
        $absText += "architect or engineer who provided this document."
        
        Write-WordText -WordSelection $selection -Text "Abstract" -Style "Heading 1" -NewLine $true
        Write-WordText -WordSelection $selection -Text $absText -NewLine $true
        
        #Write-WordText -WordSelection $selection -Text "Revision History" -Style "Heading 1" -NewLine $true
        #Write-RevisionTable
    
        Write-WordTableGrid -Caption "Revision History" -Rows 4 -ColumnHeadings ("Version","Date","Description","Author")
        
        $selection.InsertNewPage()
    
        Write-WordTableGrid -Caption "Summary of Findings" -Rows 4 -ColumnHeadings ("Item", "Explanation")
        Write-WordTableGrid -Caption "Summary of Recommendations" -Rows 4 -ColumnHeadings ("Item", "Severity", "Explanation")
    
        $selection.InsertNewPage()
    
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '1' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '2' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '3' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '4' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '5' -Doc $doc -Selection $selection -LogFile $logfile 
    
        if ($detailed -eq $true) {
            Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '5' -Detailed $true -Doc $doc -Selection $selection -LogFile $logfile 
        }
    
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '6' -Doc $doc -Selection $selection -LogFile $logfile 
    }
    else {
        Write-Log -Message "unable to load Healthcheck or Messages XML data" -Severity 3 -LogFile $logfile
        Write-Error "failed to load configuration data from XML files"
        $error.Clear()
    }
    if ($toc -ne $null) {
        $doc.TablesOfContents.Item(1).Update()
        if ($bLogValidation -eq $False) {
            Write-Host "ending healthcheck report"
            Write-Host "================="
        }
        else {
            Write-Log -Message "Ending HealthCheck Export" -LogFile $logfile
            Write-Log -Message "=================" -LogFile $logfile
        }
    }
    $time2   = Get-Date -Format "hh:mm:ss"
    $RunTime = New-TimeSpan $time1 $time2
    $Difference = "{0:g}" -f $RunTime
    Write-Output "completed in (HH:MM:SS) $Difference"
    Stop-Transcript
}
Export-ModuleMember -Function Export-CMHealthcheck
