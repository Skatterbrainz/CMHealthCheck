function Export-CMHealthReport {
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
        [parameter (Mandatory = $False, HelpMessage = "Report output type (HTML or MS Word)")]
            [ValidateSet('HTML','Word')]
            [string] $ReportType = 'Word',
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
            $Healthcheckdebug = $False
    )
    switch ($ReportType) {
        'HTML' {
            if ($AutoConfig) {
                Export-CMHealthCheckHTML -ReportFolder $ReportFolder -AutoConfig -CustomerName $CustomerName -CopyrightName $CopyrightName -Detailed -Overwrite
            }
            else {
                Export-CMHealthCheckHTML -ReportFolder $ReportFolder -CustomerName $CustomerName -CopyrightName $CopyrightName -Detailed -Overwrite
            }
            break;
        }
        'Word' {
            if ($AutoConfig) {
                Export-CMHealthCheck -ReportFolder $ReportFolder -AutoConfig -CustomerName $CustomerName -CopyrightName $CopyrightName -Detailed -Overwrite
            }
            else {
                Export-CMHealthCheck -ReportFolder $ReportFolder -CustomerName $CustomerName -CopyrightName $CopyrightName -Detailed -Overwrite
            }
            break;
        }
    }
}

Export-ModuleMember -Function Export-CMHealthReport