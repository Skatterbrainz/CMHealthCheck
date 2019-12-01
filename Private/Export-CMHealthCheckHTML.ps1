function Export-CMHealthCheckHTML {
    <#
    .SYNOPSIS
        Publish CMHealthCheck HTML Report
    .DESCRIPTION
        Converts the data output from Get-CMHealthCheck to generate an HTML report file
    .PARAMETER ReportFolder
        Path to output data folder (e.g. ".\2019-03-06\cm01.contoso.local")
    .PARAMETER OutputFolder
        Path to write new report file (default = User Documents folder)
    .PARAMETER Detailed
        Collect more granular data for final reporting, or use AutoConfig file
    .PARAMETER CustomerName
        Name of customer (default = "Customer Name"), or use AutoConfig file
    .PARAMETER AutoConfig
        Use an auto configuration file, cmhealthconfig.txt in $env:USERPROFILE\documents folder
        to fill-in AuthorName, CopyrightName, Theme, CssFilename, TableRowStyle
    .PARAMETER AuthorName
        Report Author name (default = "Your Name"), or use AutoConfig file
    .PARAMETER CopyrightName
        Text to use for copyright footer string (default = "Your Company Name"), or use AutoConfig file
    .PARAMETER Theme
        CSS style theme name, or 'Custom' to specify a file (default = 'Ocean'), or use AutoConfig file
    .PARAMETER CssFilename
        CSS file path to import when Theme is set to 'Custom', or use AutoConfig file
    .PARAMETER TableRowStyle
        Apply CSS table style: Solid, Alternating, or Dynamic. Default is Solid, or use AutoConfig file
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
        Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein"
    .EXAMPLE
        Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Overwrite -Verbose
    .EXAMPLE
        Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -OutputFolder "c:\reports" -Detailed -CustomerName "Contoso" -AuthorName "David Stein" -CopyrightName "ACME Consulting" -Theme 'Ocean' -TableRowStyle Dynamic -Verbose
    .EXAMPLE
        Export-CMHealthCheckHTML -ReportFolder "2019-03-06\cm01.contoso.local" -AutoConfig -CustomerName "Contoso"
        Applies custom parameters using "cmhealthconfig.txt" file in $env:USERPROFILE\Documents folder
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
        [Parameter (Mandatory, HelpMessage = "Collected data folder")]
            [ValidateNotNullOrEmpty()]
            [string] $ReportFolder,
        [parameter(HelpMessage="Log folder path")]
            [ValidateNotNullOrEmpty()]
            [string] $OutputFolder = "$($env:USERPROFILE)\Documents",
        [Parameter (HelpMessage = "Export full data, not only summary")]
            [switch] $Detailed,
        [parameter (HelpMessage = "Customer company name")]
            [string] $CustomerName = "Customer Name",
        [parameter (HelpMessage = "Use Auto Config File")]
            [switch] $AutoConfig,
		[parameter ()] [string] $SmsProvider,
        [parameter (HelpMessage = "Author's full name")]
            [string] $AuthorName = "",
        [parameter (HelpMessage = "Footer text")]
            [string] $CopyrightName  = "Skatterbrainz",
        [Parameter (HelpMessage = "HealthCheck query file name")]
            [string] $Healthcheckfilename = "",
        [Parameter (HelpMessage = "HealthCheck messages file name")]
            [string] $MessagesFilename = "",
        [Parameter (HelpMessage = "Debug more?")]
            $Healthcheckdebug = $False,
        [parameter (HelpMessage = "Theme Name")]
            [ValidateSet('Ocean','Emerald','Monochrome','Custom')]
            [string] $Theme = 'Ocean',
        [parameter (HelpMessage = "CSS template file")]
            [string] $CssFilename = "",
        [parameter (HelpMessage = "Table Row Style option")]
            [ValidateSet('Solid','Alternating','Dynamic')]
            [string] $TableRowStyle = 'Solid',
        [parameter (HelpMessage = "Image Log file")]
            [string] $ImageFile = "",
        [parameter (HelpMessage = "Overwrite existing report file")]
            [switch] $Overwrite
    )
    $time1      = Get-Date -Format "hh:mm:ss"
    $ModuleData = Get-Module CMHealthCheck
    $ModuleVer  = $ModuleData.Version -join '.'
    $ModulePath = $ModuleData.Path -replace 'CMHealthCheck.psm1', ''
	$logFolder  = Join-Path -Path $OutputFolder -ChildPath "_Logs"
    $logfile    = Join-Path -Path $LogFolder -ChildPath "Export-CMHealthCheckHTML.log"
    $tsLog      = Join-Path -Path $LogFolder -ChildPath "Export-CMHealthCheckHTML-Transcript.log"
    if (-not (Test-Path $logFolder)) {
        Write-Verbose "creating log folder: $logFolder"
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
    }
    else {
        Write-Verbose "log folder already exists: $logFolder"
    }

    try {
        Stop-Transcript -ErrorAction SilentlyContinue
    }
    catch {}
    finally {
        Start-Transcript -Path $tsLog -Append -ErrorAction SilentlyContinue
    }
    $Script:TableRowStyle = $TableRowStyle
    $TempFilename   = "cmhealthreport.htm"
    $bLogValidation = $False
    $bAutoProps     = $True
    $poshversion    = $PSVersionTable.PSVersion.Major
    $osversion      = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
    $ReportFile     = Join-Path -Path $OutputFolder -ChildPath "cmhealthreport`-$SmsProvider-$(Get-Date -f 'yyyyMMdd').htm"

    if ($Healthcheckfilename -eq "") {
        $Healthcheckfilename = Join-Path -Path $ModulePath -ChildPath "assets\cmhealthcheck.xml"
    }
    Write-Verbose "using healthcheck file: $HealthcheckFilename"

	if ($MessagesFilename -eq "") {
        $MessagesFilename = Join-Path -Path $ModulePath -ChildPath "assets\messages.xml"
    }
    if ($ImageFile -eq "") {
        $ImageFile = Join-Path -Path $ModulePath -ChildPath "assets\cmhclogo-275x237.png"
    }
	Write-Log -Message "converting logo image to base64 encoding" -LogFile $logfile
	$ImageData = Convert-Image2Base64 -Path $ImageFile
	if ($ImageData) {
		$LogoTag = "<img src=`"$ImageData`" width=`"100`" />"
		Write-Log -Message "image tag has been updated" -LogFile $logfile
	}
	else {
		$LogoTag = ""
		Write-Log -Message "no logo image was provided" -LogFile $logfile 
	}
    Write-Verbose "using messages file: $MessagesFilename"

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
                    }
                    'False' {
                        Set-Variable -Name $rowset[0] -Value $False
                        Write-Verbose "...$($rowset[0]) == $($rowset[1])"
                    }
                    default {
                        Set-Variable -Name $rowset[0] -Value $rowset[1]
                        Write-Verbose "...$($rowset[0]) == $($rowset[1])"
                    }
                }
            }
        }
    }
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

		Write-Log -Message "geenerating HTML report file from collected data" -LogFile $logFile -ShowMsg

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
        $htmlContent += "`n<table class=`"reportTable`" style=`"border-color:#fff`">"
		$htmlContent += "`n<tr><td style=`"width:120px`">`n"
        $htmlContent += "$LogoTag"
        $htmlContent += "`n</td><td style=`"vertical-align:top`">`n<h1>CMHealthCheck Report</h1>"
        $htmlContent += "`n<p>System Center Configuration Manager Healthcheck Status Report</p>"
        $htmlContent += "</td></tr>`n</table>"

        $htmlTable = [ordered]@{
            Customer       = $CustomerName
            Author         = "$AuthorName ($($env:USERNAME))"
            ReportDate     = (Get-Date).ToLongDateString()
            ReportFolder   = $ReportFolder
            WindowsVersion = $(Get-CimInstance -ClassName Win32_OperatingSystem).Caption
            ComputerName   = $($env:COMPUTERNAME)
        }
        $htmlContent += New-HtmlTableVertical -Caption "Report Information" -TableHash $htmlTable

        Write-Log -Message "--- inserting abstract content block" -LogFile $logfile

        $htmlContent += "`n<table class=`"reportTable`"><tr><td>"
        $htmlContent += "This document provides a point-in-time report of the current state of the System Center Configuration Manager site environment for $CustomerName. "
        $htmlContent += "For questions, concerns or comments, please consult the author of this assessment report. "
        $htmlContent += "This report was generated using CMHealthCheck $ModuleVer on $(Get-Date). Thanks to Raphael Perez and David O'Brien for "
        $htmlContent += "their work which laid the foundation on which this code was developed."
        $htmlContent += "</td></tr>`n</table>"

        Write-Log -Message "--- entering section reports" -LogFile $logfile

        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '1' -LogFile $logfile
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '2' -LogFile $logfile
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '3' -LogFile $logfile
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '4' -LogFile $logfile
        if ($Detailed) {
            $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '5' -LogFile $logfile -Detailed
        }
		else {
            $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '5' -LogFile $logfile
        }
        $htmlContent += Write-HtmlReportSection -HealthCheckXML $HealthCheckXML -Section '6' -LogFile $logfile

        #Set-DocAppendix

        Write-Log -Message "inserting copyright footer" -LogFile $logfile
        $htmlContent += "`n<p class=`"footer`">CMHealthCheck $ModuleVer by David Stein. Copyright &copy; $((Get-Date).Year) $CopyrightName"
        $htmlContent += " - <a href=`"https://github.com/skatterbrainz/CMHealthCheck`" target=`"_blank`">GitHub`:CMHealthCheck</a></p>"
        $htmlContent += "`n</body></html>"

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
