#requires -version 5
function Export-CMHealthCheck {
    <#
    .SYNOPSIS
        Convert extracted ConfigMgr site information to Word Document
    .DESCRIPTION
        Converts the data output from Get-CMHealthCheck to generate a
        report document using Microsoft Word (2010, 2013, 2016). Intended
        to be invoked on a desktop computer which has Office installed.
    .PARAMETER ReportFolder
        Path to output data folder (e.g. ".\2019-03-06\cm01.contoso.local")
    .PARAMETER AutoConfig
        Use an auto configuration file, cmhealthconfig.txt in $env:USERPROFILE\documents folder
        to fill-in AuthorName, CopyrightName, Theme, CssFilename, TableRowStyle
    .PARAMETER Detailed
        Collect more granular data for final reporting, or use AutoConfig file
    .PARAMETER CoverPage
        Word theme cover page (default = "Slice (Light)"), or use AutoConfig file
    .PARAMETER Template
        Word document file to use as a template. Should have a cover page already in place.
        If Template is specified, CoverPage and Copyright are ignored.
    .PARAMETER CustomerName
        Name of customer (default = "Customer Name"), or use AutoConfig file
    .PARAMETER AuthorName
        Report Author name (default = "Your Name"), or use AutoConfig file
    .PARAMETER CopyrightName
        Text to use for copyright footer string (default = "Your Company Name")
    .PARAMETER Overwrite
        Overwrite existing report file if found, or use AutoConfig file
    .PARAMETER Healthcheckfilename
        Healthcheck configuration XML file name (default = ".\assets\cmhealthcheck.xml")
        The file can be local, UNC or URI sourced as well
    .PARAMETER MessagesFilename
        Status and error message lookup table (default = ".\assets\messages.xml")
        The file can be local, UNC or URI sourced as well
    .PARAMETER Healthcheckdebug
        Enable verbose output (or use -Verbose)
    .EXAMPLE
        Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
    .EXAMPLE
        Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -Template ".\contoso.docx" -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
    .EXAMPLE
        Export-CMHealthCheck -ReportFolder "2019-03-06\cm01.contoso.local" -AutoConfig -CustomerName "Contoso"
    .NOTES
        Sample AutoConfig file cmhealthconfig.txt...
        AuthorName=John Wick
        CopyrightName=Retirement Specialists
        Theme=Ocean
        Detailed=True
        TableRowStyle=Solid
        CssFilename=c:\docs\wickrocks.css
        ImageFile=c:\docs\bodybags.png
        CoverPage=
        Template=
        HealthcheckFilename=
        MessagesFilename=
        HealthcheckDebug=False
        Overwrite=True
    #>
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $True, HelpMessage = "Collected data folder")] 
            [ValidateNotNullOrEmpty()]
            [string] $ReportFolder,
        [parameter(Mandatory=$False, HelpMessage="Log folder path")]
            [ValidateNotNullOrEmpty()]
            [string] $OutputFolder = "$($env:USERPROFILE)\Documents",
        [parameter (Mandatory = $False, HelpMessage = "Customer company name")] 
            [string] $CustomerName = "Customer Name",
        [parameter (Mandatory = $False, HelpMessage = "Use Auto Config File")]
            [switch] $AutoConfig,
        [Parameter (Mandatory = $False, HelpMessage = "Export full data, not only summary")] 
            [switch] $Detailed,
        [parameter (Mandatory = $False, HelpMessage = "Word Template cover page name")] 
            [string] $CoverPage = "Slice (Light)",
        [parameter (Mandatory = $False, HelpMessage = "Word document source file")]
            [string] $Template = "", 
        [parameter (Mandatory = $False, HelpMessage = "Author's full name")] 
            [string] $AuthorName = "Your Name",
        [parameter (Mandatory = $False, HelpMessage = "Footer text")]
            [string] $CopyrightName  = "Your Company Name",
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck query file name")] 
            [string] $Healthcheckfilename = "", 
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck messages file name")]
            [string] $MessagesFilename = "", 
        [Parameter (Mandatory = $False, HelpMessage = "Debug more?")] 
            $Healthcheckdebug = $False,
        [parameter (Mandatory = $False, HelpMessage = "Overwrite existing report file")]
            [switch] $Overwrite
    )
    $time1 = Get-Date -Format "hh:mm:ss"
    $ModuleData = Get-Module CMHealthCheck
    $ModuleVer  = $ModuleData.Version -join '.'
    $ModulePath = $ModuleData.Path -replace 'CMHealthCheck.psm1', ''
    $tsLog      = Join-Path -Path $OutputFolder -ChildPath "Export-CMHealthCheck-Transcript.log"
    $logfile    = Join-Path -Path $OutputFolder -ChildPath "Export-CMHealthCheck.log"
    try {
        Stop-Transcript -ErrorAction SilentlyContinue
    }
    catch {}
    finally {
        Start-Transcript -Path $tsLog -Append -ErrorAction SilentlyContinue
    }
    
    $TempFilename      = "cmhealthreport.docx"
    $DefaultTableStyle = "Grid Table 4 - Accent 1"
    $TableStyle        = "Grid Table 4 - Accent 1"
    $TableSimpleStyle  = "Grid Table 4 - Accent 1"
    $ReviewTableStyle  = "Grid Table 4 - Accent 6"
    $RecTableStyle     = "Grid Table 4 - Accent 3"
    $ReviewTableCols   = ("No.", "Severity", "Comment")
    $bLogValidation    = $False
    $bAutoProps        = $True
    $NormalFontSize    = 10
    $poshversion       = $PSVersionTable.PSVersion.Major
    $osversion         = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    #$FormatEnumerationLimit = -1

    $autoconfigfile = Join-Path -Path $env:USERPROFILE -ChildPath "documents\cmhealthconfig.txt"
    if ($AutoConfig -and (Test-Path $autoconfigfile)) {
        Write-Verbose "importing settings from config file: $autoconfigfile"
        $cfgdata = Get-Content -Path $autoconfigfile
        $cfgdata | % {
            $rowset = $_ -split '='
            if (![string]::IsNullOrEmpty($rowset[1])) {
                switch($rowset[1]) {
                    'True' {
                        Set-Variable -Name $rowset[0] -Value $True
                        Write-Verbose "...$($rowset[0]) == $($rowset[1])"
                        break
                    }
                    'False' {
                        Set-Variable -Name $rowset[0] -Value $False
                        Write-Verbose "...$($rowset[0]) == $($rowset[1])"
                        break
                    }
                    default {
                        Set-Variable -Name $rowset[0] -Value $rowset[1]
                        Write-Verbose "...$($rowset[0]) == $($rowset[1])"
                        break
                    }
                }
            }
        }
    }

    if ($Healthcheckfilename -eq "") {
        $Healthcheckfilename = Join-Path -Path $ModulePath -ChildPath "assets\cmhealthcheck.xml"
    }
	if ($MessagesFilename -eq "") {
        $MessagesFilename = Join-Path -Path $ModulePath -ChildPath "assets\messages.xml"
    }

    if ($healthcheckdebug -eq $true) { 
        $PSDefaultParameterValues = @{"*:Verbose"=$True}
    }
    $logFolder = Join-Path -Path $PWD.Path -ChildPath "_Logs\"
    if (-not (Test-Path $logFolder)) {
        mkdir $logFolder -Force | Out-Null
    }
    if ($reportFolder.Substring($reportFolder.length-1) -ne '\') { $reportFolder+= '\' }
    
    $Error.Clear()

    $poshversion = $PSVersionTable.PSVersion.Major
    Show-CMHCInfo
    
    [xml]$HealthCheckXML = Get-CmHealthCheckFile -XmlSource $HealthcheckFilename
    [xml]$MessagesXML    = Get-CmHealthCheckFile -XmlSource $MessagesFilename
     
    Write-Log -Message "Connecting to Microsoft Word..." -LogFile $logfile
    try {
        $Word = New-Object -ComObject "Word.Application" -ErrorAction Stop
    }
    catch {
        Invoke-Error -Message "Microsoft Word could not be opened!"
        break
    }

    if ($HealthCheckXML -and $MessagesXML) {
        $bLogValidation = $true
        Write-Log -Message "Provisioning config table" -LogFile $logfile
        $ConfigTable = New-Object System.Data.DataTable 'ConfigTable'
        $ConfigTable = Get-CmXMLFile -Path $reportFolder -FileName "config.xml"
        if ($ConfigTable -eq "") {
            Invoke-Error -Message "File $configfile does not exist, no futher action taken"; break
        }
        Write-Log -Message "Provisioning report table" -LogFile $logfile
        $ReportTable = New-Object System.Data.DataTable 'ReportTable'
        $ReportTable = Get-CmXMLFile -Path $reportFolder -FileName "report.xml"
        if ($ReportTable -eq "") {
            Invoke-Error -Message "File $repfile does not exist, no futher action taken"; break
        }
        Write-Log -Message "Assigning number of days from config data..." -LogFile $logfile
        if ($poshversion -eq 3) { 
            $NumberOfDays = $ConfigTable.Rows[0].NumberOfDays
        }
        else { 
            $NumberOfDays = $ConfigTable.NumberOfDays
        }

        if (!(Test-Powershell64bit)) { Invoke-Error -Message "Powershell is not 64bit, no futher action taken"; break }
        
        $wordVersion = $Word.Version
        Write-Log -Message "Microsoft Word version: $WordVersion" -LogFile $logfile
        if ($wordVersion -lt '15.0') { Invoke-Error -Message "This module requires Word 2013 or newer"; break }
        if ($Template -ne "") {
            $newFile = Get-WordTempSource -SourceFile $Template
            Write-Log -Message "Opening temp file [$newFile]..." -LogFile $logfile
            try {
                $Doc = $Word.Documents.Open($newFile)
            }
            catch {
                Invoke-Error -Message "Failed to open temp document file: $newFile"
                break
            }
        }
        else {
            Write-Log -Message "Creating new (blank) document..." -LogFile $logfile
            $Doc = $Word.Documents.Add()
        }
        if ($doc -eq $null) { Invoke-Error -Message "Failed to obtain handle to Word document"; break }
        $Word.Visible = $True
        $Selection = $Word.Selection

        Set-WordOptions
        Set-DocProperties

        Write-Log -Message "Loading default building blocks " -LogFile $logfile
        $Word.Templates.LoadBuildingBlocks() | Out-Null	
        $BuildingBlocks = $Word.Templates | Where-Object {$_.name -eq "Built-In Building Blocks.dotx"}

        if ($Template -eq "") {
            Write-Log -Message "Inserting cover page: $CoverPage" -LogFile $logfile
            $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
            $part.Insert($selection.Range,$True) | Out-Null
        }
        else {
            Write-Log -Message "Cover page option ignored when using custom template" -LogFile $logfile
            $Selection.EndKey(6, 0) | Out-Null            
        }

        $selection.InsertNewPage()
        Set-WordTOC

        $selection.InsertNewPage()
        $currentview = $doc.ActiveWindow.ActivePane.view.SeekView
        $doc.ActiveWindow.ActivePane.view.SeekView = 4
        Set-WordFooter

        $doc.ActiveWindow.ActivePane.view.SeekView = $currentview
        $selection.EndKey(6,0) | Out-Null
        Set-WordAbstract

        Write-WordTableGrid -Caption "Revision History" -Rows 4 -ColumnHeadings ("Version","Date","Description","Author")

        $selection.InsertNewPage()
        Write-WordTableGrid -Caption "Summary of Findings" -Rows 4 -ColumnHeadings ("Item", "Severity", "Explanation") -StyleName $ReviewTableStyle
        Write-WordTableGrid -Caption "Summary of Recommendations" -Rows 4 -ColumnHeadings ("Item", "Severity", "Explanation") -StyleName $RecTableStyle

        $selection.InsertNewPage()
        Write-DocReportSections

        $selection.InsertNewPage()
        Set-DocAppendix
    }
    else {
        Write-Log -Message "Unable to load Healthcheck or Messages XML data" -Severity 3 -LogFile $logfile -ShowMsg
        $error.Clear()
    }

    if ($toc -ne $null) {
        $Doc.TablesOfContents.Item(1).Update()
        if ($bLogValidation -eq $False) {
            Write-Host "Finishing up healthcheck report"
        }
        else {
            Write-Log -Message "Finishing up HealthCheck Export" -LogFile $logfile
        }
    }
    $Difference = Get-TimeOffset -StartTime $time1
    Write-Log -Message "Completed in: $Difference (hh:mm:ss)" -LogFile $logfile -ShowMsg
    Stop-Transcript
}
