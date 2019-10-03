<#
.SYNOPSIS
	Generate Health Information from a Configuration Manager site
.DESCRIPTION
	Generate Health Information from a Configuration Manager site
.PARAMETER SmsProvider
	FQDN of the SMS Provider host in the Configuration Manager site
.PARAMETER OutputFolder
    Path to output data during collection phase
.PARAMETER PublishFolder
    Path to save the HTML report file
.PARAMETER CustomerName
    Name of customer (default = "Customer Name"), or use AutoConfig file
.PARAMETER AuthorName
    Report Author name (default = "Your Name"), or use AutoConfig file
.PARAMETER CopyrightName
    Text to use for copyright footer string (default = "Your Company Name")
.PARAMETER MessagesFilename
    Status and error message lookup table (default = ".\assets\messages.xml")
    The file can be local, UNC or URI sourced as well
.PARAMETER HealthcheckFilename
    Name of configuration file (default is .\assets\cmhealthcheck.xml)
.PARAMETER Healthcheckdebug
    Enable verbose output (or use -Verbose)
.PARAMETER NumberOfDays
    Number of days to go back for alerts in logs (default = 7)
.PARAMETER Overwrite
    Overwrite existing output folder if found.
    Folder is named by datestamp, so this only applies when
    running repeatedly on the same date
.PARAMETER NoHotFix
    Skip inventory of installed hotfixes
.PARAMETER OpenBrowser
    Open HTML report in default web browser upon completion
.PARAMETER AutoConfig
    Load parameters from configuration file
    Example:
        ```
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
        ```

.PARAMETER Detailed
    Display additional details (verbose)
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp"
    Standard/default settings to collect data, and generate HTML report
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -Overwrite
    Replaces an existing (previous) output from the same date
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -OpenBrowser
    Opens the HTML report in the default web browser, upon completion
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -Detailed
    Generates additional detail in the output report file
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -AutoConfig "config.txt"
    Loads reporting parameters from custom text file
.EXAMPLE
    Invoke-CMHealthCheck -SmsProvider "cm01.contoso.local" -CustomerName "Contoso" -AuthorName "Skatter Brainz" -CopyrightName "SkatterCorp" -NoHotFix
    Skips inventory of installed operating system hotfixes
.NOTES
    New function for 1.0.11 - 10/2019
.OUTPUTS
    collected data files, folders, and output HTML file, happiness, confusion, consternation, whatever that means
.LINK
    https://github.com/Skatterbrainz/CMHealthCheck/blob/master/Docs/Invoke-CMHealthCheck.md
#>
function Invoke-CMHealthCheck {
    [CmdletBinding(ConfirmImpact="Low")]
    param (
        [parameter()] [ValidateNotNullOrEmpty()][string] $SmsProvider = $($env:COMPUTERNAME),
        [parameter()] [string] $CustomerName = "Customer Name",
        [parameter()] [string] $AuthorName = "Your Name",
        [parameter()] [string] $CopyrightName  = "Your Company Name",
        [parameter()] [string] $OutputFolder = "$($env:USERPROFILE)\Documents",
        [parameter()] [ValidateNotNullOrEmpty()] [string] $PublishFolder = "$($env:USERPROFILE)\Documents",
        [parameter()] [int] $NumberOfDays = 7,
		[parameter()] [string]$ReportType = 'HTML',
		[parameter()] [switch] $OpenBrowser,
        [parameter()] [switch] $NoHotfix ,
        [parameter()] [switch] $OverWrite,
        [parameter()] [switch] $AutoConfig,
        [parameter()] [switch] $Detailed,
        [parameter()] [switch] $Healthcheckdebug,
        [parameter()] [string] $Healthcheckfilename = "",
        [parameter()] [string] $MessagesFilename = ""
    )
    
    $ReportFolder = Join-Path $OutputFolder "$(Get-Date -f 'yyyy-MM-dd')\$SmsProvider"
    try {
        Write-Verbose "report folder path = $ReportFolder"
        $getParams = @{
            SmsProvider   = $SmsProvider
            OutputFolder  = $OutputFolder
            NumberOfDays  = $NumberOfDays
            NoHotfix      = $NoHotfix
			OverWrite     = $OverWrite
            Verbose       = $VerbosePreference
        }
        Write-Verbose "calling Get-CMHealthCheck with parameter set"
        Get-CMHealthCheck @getParams
    }
    catch {
        Write-Error $_.Exception.Message
    }
	Write-Log "------------------ begin report publishing ---------------------"
    try {
        if (Test-Path $ReportFolder) {
            Write-Verbose "calling Export-CMHealthCheck with parameter set"
            $expParams = @{
                ReportType       = $ReportType
                ReportFolder     = $ReportFolder
                OutputFolder     = $OutputFolder
                CustomerName     = $CustomerName
                AutoConfig       = $AutoConfig
                Detailed         = $Detailed
                CoverPage        = $CoverPage
                Template         = $Template
                AuthorName       = $AuthorName
                CopyrightName    = $CopyrightName
                MessagesFilename = $MessagesFilename
                Healthcheckdebug = $Healthcheckdebug
                Healthcheckfilename = $Healthcheckfilename
                Verbose          = $VerbosePreference
            }
            Export-CMHealthReport @expParams
			if ($OpenBrowser) {
				$newFile = Join-Path -Path $OutputFolder -ChildPath "cmhealthreport`-$SmsProvider-$(Get-Date -f 'yyyyMMdd').htm"
				if (Test-Path $newFile) {
					Write-Host "opening report in default web browser: $newFile" -ForegroundColor Cyan
					Start-Process $newFile
				}
				else {
					Write-Warning "file not found: $newFile"
				}
			}
        }
        else {
            throw "report folder not found: $ReportFolder"
        }
    }
    catch {
        Write-Error $_.Exception.Message
    }
}