#requires -version 3
function Export-CMHealthCheck {
    <#
    .SYNOPSIS
        Convert extracted ConfigMgr site information to Word Document
    .DESCRIPTION
        Converts the data output from Get-CMHealthCheck to generate a
        report document using Microsoft Word (2010, 2013, 2016). Intended
        to be invoked on a desktop computer which has Office installed.
    .PARAMETER ReportFolder
        Path to output data folder (e.g. ".\2017-11-17\cm01.contoso.com")
    .PARAMETER Detailed
        Collect more granular data for final reporting
    .PARAMETER CoverPage
        Word theme cover page (default = "Slice (Light)")
    .PARAMETER CustomerName
        Name of customer (default = "Customer Name")
    .PARAMETER AuthorName
        Report Author name (default = "Your Name")
    .PARAMETER CopyrightName
        Text to use for copyright footer string (default = "Your Company Name")
    .PARAMETER Overwrite
        Overwrite existing report file if found
    .PARAMETER Healthcheckfilename
        Healthcheck configuration XML file name (default = ".\assets\cmhealthcheck.xml")
        The file can be local, UNC or URI sourced as well
    .PARAMETER MessagesFilename
        Status and error message lookup table (default = ".\assets\messages.xml")
        The file can be local, UNC or URI sourced as well
    .PARAMETER Healthcheckdebug
        Enable verbose output (or use -Verbose)
    .EXAMPLE
        Export-CMHealthCheck -ReportFolder "2017-11-17\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
    .NOTES
        1.0.1 - 11/17/2017 - David Stein
        Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
        Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
        Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
        Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference
    #>
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
            [string] $Healthcheckfilename = "", #'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/cmhealthcheck.xml',
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck messages file name")]
            [string] $MessagesFilename = "", #'https://raw.githubusercontent.com/Skatterbrainz/CM_HealthCheck/master/Messages.xml',
        [Parameter (Mandatory = $False, HelpMessage = "Debug more?")] 
            $Healthcheckdebug = $False,
        [parameter (Mandatory = $False, HelpMessage = "Overwrite existing report file")]
            [switch] $Overwrite
    )
    $time1 = Get-Date -Format "hh:mm:ss"
    Start-Transcript -Path ".\_logs\export-reportfile.log" -Append
    $bLogValidation = $False
    $bAutoProps     = $True
    $NormalFontSize = 10
    $poshversion    = $PSVersionTable.PSVersion.Major
    $osversion      = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    $FormatEnumerationLimit = -1
    
    if ($Healthcheckfilename -eq "") {
        $ModulePath = $((Get-Module CMHealthCheck).Path -replace ('CMHealthCheck.psm1', ''))
        $Healthcheckfilename = "$ModulePath"+"assets\cmhealthcheck.xml"
    }
	if ($MessagesFilename -eq "") {
        $ModulePath = $((Get-Module CMHealthCheck).Path -replace ('CMHealthCheck.psm1', ''))
        $MessagesFilename = "$ModulePath"+"assets\messages.xml"
    }

    if ($healthcheckdebug -eq $true) { $PSDefaultParameterValues = @{"*:Verbose"=$True}; $currentFolder = "C:\Temp\CMHealthCheck\" }
    $logFolder = $PWD.Path + "\_Logs\"
    if (-not (Test-Path $logFolder)) {
        Write-Error "$logFolder was not found!"
        break
    }
    if ($reportFolder.substring($reportFolder.length-1) -ne '\') { $reportFolder+= '\' }
    $component = ($MyInvocation.MyCommand.Name -replace '.ps1', '')
    $logfile = $logFolder + $component + ".log"
    $Error.Clear()

    Write-Log -Message "==========" -LogFile $logfile -ShowMsg $false
    Write-Log -Message "Script Version...: $($Script:ScriptVersion)" -LogFile $logfile
    Write-Log -Message "Windows Version..: $osversion" -LogFile $logfile
    Write-Log -Message "environment......: Running Powershell version: $poshversion" -LogFile $logfile
    Write-Log -Message "environment......: Running Powershell 64 bits" -LogFile $logfile
    Write-Log -Message "Report Folder....: $reportFolder" -LogFile $logfile
    Write-Log -Message "Detailed Report..: $detailed" -LogFile $logfile
    Write-Verbose "Export-CMHealthCheck $ScriptVersion"
    Write-Verbose "Current Folder..: $($PWD.Path)"
    Write-Verbose "Log Folder......: $logFolder"
    Write-Verbose "Log File........: $logfile" 
    Write-Verbose "control file....: $Healthcheckfilename"
    Write-Verbose "message file....: $MessagesFilename"

    Write-Host "Export-CMHealthCheck - $ScriptVersion - https://github.com/Skatterbrainz/CMHealthCheck" -ForegroundColor Green
    $poshversion = $PSVersionTable.PSVersion.Major
    
    [xml]$HealthCheckXML = Get-CmHealthCheckFile -XmlSource $HealthcheckFilename
    [xml]$MessagesXML    = Get-CmHealthCheckFile -XmlSource $MessagesFilename
     
    Write-Verbose "info: connecting to Microsoft Word..."
    try {
        $Word = New-Object -ComObject "Word.Application" -ErrorAction Stop
    }
    catch {
        Write-Warning "Error: This script requires Microsoft Word"
        Stop-Transcript -ErrorAction SilentlyContinue
        break
    }

    if ($HealthCheckXML -and $MessagesXML) {
        if (Test-Folder -Path $logFolder) {
            try {
                New-Item ($logFolder + 'Test.log') -Type File -Force | Out-Null 
                Remove-Item ($logFolder + 'Test.log') -Force | Out-Null 
            }
            catch {
                Write-Warning "Unable to read/write file on $logFolder folder, no futher action taken"
                Stop-Transcript -ErrorAction SilentlyContinue
                break    
            }
        }
        else {
            Write-Host "Unable to create Log Folder, no futher action taken" -ForegroundColor Red
            Stop-Transcript -ErrorAction SilentlyContinue
            break
        }
        $bLogValidation = $true
    
        if (Test-Folder -Path $reportFolder -Create $false) {
            if (!(Test-Path -Path ($reportFolder + "config.xml"))) {
                Write-Log -Message "File $($reportFolder)config.xml does not exist, no futher action taken" -Severity 3 -LogFile $logfile
                Stop-Transcript -ErrorAction SilentlyContinue
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
                Stop-Transcript -ErrorAction SilentlyContinue
                break
            }
            else {
                $ReportTable = New-Object System.Data.DataTable 'ReportTable'
                $ReportTable = Import-CliXml -Path ($reportFolder + "report.xml")
            }
        }
        else {
            Write-Warning "Folder: $reportFolder does not exist, no futher action taken"
            Stop-Transcript -ErrorAction SilentlyContinue
            break
        }
        
        if (!(Test-Powershell64bit)) {
            Write-Log -Message "Powershell is not 64bit, no futher action taken" -Severity 3 -LogFile $logfile
            Stop-Transcript -ErrorAction SilentlyContinue
            break
        }
        
        $wordVersion = $Word.Version
        Write-Log -Message "Word Version: $WordVersion" -LogFile $logfile	
        Write-Verbose "info: Microsoft Word version: $WordVersion"
        $styles = Set-WordFormatting
        if ($styles) {
            $TableStyle = $styles[0]
            $TableSimpleStyle = $styles[1]
        }
        else { 
            Write-Log -Message "This script requires Word 2010 to 2016 version, no further action taken" -Severity 3 -LogFile $logfile 
            Stop-Transcript -ErrorAction SilentlyContinue
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
            Stop-Transcript -ErrorAction SilentlyContinue
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
            
        Write-WordTableGrid -Caption "Revision History" -Rows 4 -ColumnHeadings ("Version","Date","Description","Author")
        
        $selection.InsertNewPage()
    
        Write-WordTableGrid -Caption "Summary of Findings" -Rows 4 -ColumnHeadings ("Item", "Explanation")
        Write-WordTableGrid -Caption "Summary of Recommendations" -Rows 4 -ColumnHeadings ("Item", "Severity", "Explanation")
    
        $selection.InsertNewPage()
    
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '1' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '2' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '3' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '4' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '5' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        if ($detailed -eq $true) {
            Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '5' -Detailed $true -Doc $doc -Selection $selection -LogFile $logfile 
            Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
        }
        Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '6' -Doc $doc -Selection $selection -LogFile $logfile 
        Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings ("No.", "Comment")
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
    Write-Output "completed in: $Difference (hh:mm:ss)"
    Stop-Transcript
}
Export-ModuleMember -Function Export-CMHealthcheck