#requires -version 3
function Export-CMHealthCheckHTML {
    <#
    .SYNOPSIS
        Convert extracted ConfigMgr site information to HTML report
    .DESCRIPTION
        Converts the data output from Get-CMHealthCheck to generate an HTML report file
    .PARAMETER ReportFolder
        Path to output data folder (e.g. ".\2018-9-19\cm01.contoso.com")
    .PARAMETER Detailed
        Collect more granular data for final reporting
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
        Export-CMHealthCheckHTML -ReportFolder "2018-9-19\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein"
    .EXAMPLE
        Export-CMHealthCheckHTML -ReportFolder "2018-9-19\cm01.contoso.com" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
    .NOTES
        * 1.0.4 - 12/04/2017 - David Stein
        * 1.0.5 - 09/19/2018 - David Stein
        * Thanks to Rafael Perez for inventing this - http://www.rflsystems.co.uk
        * Thanks to Carl Webster for the basis of Word functions - http://www.carlwebster.com
        * Thanks to David O'Brien for additional Word function - http://www.david-obrien.net/2013/06/20/huge-powershell-inventory-script-for-configmgr-2012/
        * Thanks to Starbucks for empowering me to survive hours of clicking through the Office Word API reference
    #>
    [CmdletBinding()]
    param (
        [Parameter (Mandatory = $True, HelpMessage = "Collected data folder")] 
            [ValidateNotNullOrEmpty()]
            [string] $ReportFolder,
        [parameter(Mandatory=$False, HelpMessage="Log folder path")]
            [ValidateNotNullOrEmpty()]
            [string] $OutputFolder = "$($env:USERPROFILE)\Documents",
        [Parameter (Mandatory = $False, HelpMessage = "Export full data, not only summary")] 
            [switch] $Detailed,
        [parameter (Mandatory = $False, HelpMessage = "Customer company name")] 
            [string] $CustomerName = "Customer Name",
        [parameter (Mandatory = $False, HelpMessage = "Author's full name")] 
            [string] $AuthorName = "Your Name",
        [parameter (Mandatory = $False, HelpMessage = "Footer text")]
            [string] $CopyrightName  = "Skatterbrainz",
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck query file name")] 
            [string] $Healthcheckfilename = "", 
        [Parameter (Mandatory = $False, HelpMessage = "HealthCheck messages file name")]
            [string] $MessagesFilename = "", 
        [Parameter (Mandatory = $False, HelpMessage = "Debug more?")] 
            $Healthcheckdebug = $False,
        [parameter (Mandatory = $False, HelpMessage = "Theme Name")]
            [ValidateSet('Ocean','Monochrome','Custom')]
            [string] $Theme = 'Ocean',
        [parameter (Mandatory = $False, HelpMessage = "CSS template file")]
            [string] $CssFilename = "",
        [parameter (Mandatory = $False, HelpMessage = "Dynamic Table Row Styles")]
            [switch] $DynamicTableRows,
        [parameter (Mandatory = $False, HelpMessage = "Overwrite existing report file")]
            [switch] $Overwrite
    )
    $time1      = Get-Date -Format "hh:mm:ss"
    $ModuleData = Get-Module CMHealthCheck
    $ModuleVer  = $ModuleData.Version -join '.'
    $ModulePath = $ModuleData.Path -replace 'CMHealthCheck.psm1', ''
    $tsLog      = Join-Path -Path $OutputFolder -ChildPath "Export-CMHealthCheckHTML-Transcript.log"
    $logfile    = Join-Path -Path $OutputFolder -ChildPath "Export-CMHealthCheckHTML.log"
    try {
        Stop-Transcript -ErrorAction SilentlyContinue
    }
    catch {}
    finally {
        Start-Transcript -Path $tsLog -Append -ErrorAction SilentlyContinue
    }
    
    $TempFilename      = "cmhealthreport.htm"
    $ReviewTableCols   = ("No.", "Severity", "Comment")
    $bLogValidation    = $False
    $bAutoProps        = $True
    $poshversion       = $PSVersionTable.PSVersion.Major
    $osversion         = (Get-WmiObject -Class Win32_OperatingSystem).Caption
    #$FormatEnumerationLimit = -1
    
    if ($Healthcheckfilename -eq "") {
        $Healthcheckfilename = Join-Path -Path $ModulePath -ChildPath "assets\cmhealthcheck.xml"
    }
    Write-Verbose "using healthcheck file: $HealthcheckFilename"

	if ($MessagesFilename -eq "") {
        $MessagesFilename = Join-Path -Path $ModulePath -ChildPath "assets\messages.xml"
    }
    Write-Verbose "using messages file: $MessagesFilename"

    if ($Theme -eq 'Custom') {
        if ($CssFilename -eq "") {
            Write-Warning "No stylesheet was specified for [custom] option. Using default.css"
            $CssFilename = Join-Path -Path $ModulePath -ChildPath "assets\default.css"
        }
        else {
            if (!(Test-Path $CssFilename)) {
                Write-Warning "$CssFilename was not found!"
                break
            }
        }
    }
    else {
        $CssFilename = Join-Path -Path $ModulePath -ChildPath "assets\$Theme.css"
        if (!(Test-Path $CssFilename)) {
            Write-Warning "$CssFilename was not found!"
            break
        }
    }
    Write-Log -Message "importing css template: $CssFilename" -LogFile $logfile
    $css = Get-Content $CssFilename

    if ($healthcheckdebug -eq $true) { 
        $PSDefaultParameterValues = @{"*:Verbose"=$True}
    }

    $logFolder = Join-Path -Path $PWD.Path -ChildPath "_Logs\"
    $reportFile = Join-Path -Path $OutputFolder -ChildPath $TempFilename
    if (-not (Test-Path $logFolder)) {
        Write-Verbose "creating log folder: $logFolder"
        mkdir $logFolder -Force | Out-Null
    }
    else {
        Write-Verbose "log folder already exists: $logFolder"
    }
    if ($reportFolder.Substring($reportFolder.length-1) -ne '\') { $reportFolder+= '\' }
    
    $Error.Clear()

    $poshversion = $PSVersionTable.PSVersion.Major
    Show-CMHCInfo
    
    [xml]$HealthCheckXML = Get-CmHealthCheckFile -XmlSource $HealthcheckFilename
    [xml]$MessagesXML    = Get-CmHealthCheckFile -XmlSource $MessagesFilename
     
    if ($HealthCheckXML -and $MessagesXML) {
        $bLogValidation = $true
        Write-Log -Message "----- Provisioning config table -----" -LogFile $logfile
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
        
        Write-Log -Message "initializing HTML content" -LogFile $logfile
        $htmlContent = @"
<html>
    <head>
        <title="CMHealthCheck Report">
        <style type="text/css">
        $css
        </style>
        <meta content="text/html; charset=UTF-8" http-equiv="Content-Type" />
    </head>
    <body>
"@

        Write-Log -Message "inserting title caption" -LogFile $logfile
        $htmlContent += "<h1>CMHealthCheck Report</h1>"

        $htmlTable = @{
            Customer       = $CustomerName
            Author         = "$AuthorName ($($env:USERNAME))"
            ReportFolder   = $ReportFolder
            WindowsVersion = $(Get-WmiObject -Class Win32_OperatingSystem).Caption
            ReportDate     = (Get-Date).ToLongDateString()
            ComputerName   = $($env:COMPUTERNAME)
        }
        $htmlContent += New-HtmlTableVertical -Caption "Report Generation" -TableStyle "background:#eee;width:1000px" -ColumnStyle1 "background:#c0c0c0" -ColumnStyle2 "background:#eee" -TableHash $htmlTable
        
        Write-Log -Message "--- inserting abstract content block" -LogFile $logfile

        $htmlContent += "<table style=`"width:1000px`"><tr><td>"
        $htmlContent += "This document provides a point-in-time report of the current state of the System Center Configuration Manager site environment for $CustomerName. "
        $htmlContent += "For questions, concerns or comments, please consult the author of this assessment report. "
        $htmlContent += "This report was generated using CMHealthCheck $ModuleVer on $(Get-Date)."
        $htmlContent += "</td></tr></table>"

        Write-Log -Message "--- inserting revision table" -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Revision History" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Version=100,Date=100,Description" -RowStyle2 "background:#eee" -Rows 3
        Write-Log -Message "--- inserting summary findings table" -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Summary of Findings" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2
        Write-Log -Message "--- inserting recommendations table" -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Summary of Recommendations" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2

        Write-Log -Message "--- entering section reports" -LogFile $logfile

        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '1' -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '2' -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '3' -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '4' -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2

        if ($Detailed) {
            $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '5' -LogFile $logfile
            $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2
        }

        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '6' -LogFile $logfile
        $htmlContent += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Explanation" -RowStyle2 "background:#eee" -Rows 2

        #Set-DocAppendix

        Write-Log -Message "inserting copyright footer" -LogFile $logfile
        $htmlContent += "<p class=`"footer`">CMHealthCheck $ModuleVer . Copyright &copy; 2018 Skatterbrainz</p>"
        $htmlContent += "</body></html>"
        Write-Log -Message "writing output file: $ReportFile" -LogFile $logfile
        $htmlContent | Out-File -FilePath $ReportFile -Force
    }
    else {
        Write-Log -Message "Unable to load Healthcheck or Messages XML data" -Severity 3 -LogFile $logfile -ShowMsg
        $error.Clear()
    }

    $Difference = Get-TimeOffset -StartTime $time1
    Write-Log -Message "Completed in: $Difference (hh:mm:ss)" -LogFile $logfile -ShowMsg
    Stop-Transcript
}
Export-ModuleMember -Function Export-CMHealthcheckHTML
