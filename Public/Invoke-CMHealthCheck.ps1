<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER SmsProvider
.PARAMETER CustomerName
.PARAMETER AuthorName
.PARAMETER CopyrightName
.PARAMETER OutputFolder
.PARAMETER PublishFolder
.PARAMETER NumberOfDays
.PARAMETER NoHotFix
.PARAMETER OverWrite
.PARAMETER AutoConfig
.PARAMETER Detailed
.PARAMETER Healthcheckdebug
.PARAMETER Healthcheckfilename
.PARAMETER MessagesFilename
.EXAMPLE
.EXAMPLE
.NOTES
.OUTPUTS
    collected data files, folders, and output HTML file
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
        [parameter()] [bool] $NoHotfix = $False,
        [parameter()] [bool] $OverWrite = $False,
        [parameter()] [bool] $AutoConfig = $False,
        [parameter()] [bool] $Detailed = $False,
        [parameter()] [bool] $Healthcheckdebug = $False,
        [parameter()] [string] $Healthcheckfilename = "",
        [parameter()] [string] $MessagesFilename = ""
    )
    [string]$ReportType = 'HTML'
    $ReportFolder = Join-Path $OutputFolder "$(Get-Date -f 'yyyy-MM-dd')\$SmsProvider"
    try {
        Write-Verbose "report folder path = $ReportFolder"
        $getParams = @{
            SmsProvider   = $SmsProvider
            CustomerName  = $CustomerName
            AuthorName    = $AuthorName
            CopyrightName = $CopyrightName
            OutputFolder  = $OutputFolder
            NumberOfDays  = $NumberOfDays
            NoHotfix      = $NoHotfix
            Verbose       = $VerbosePreference
        }
        Write-Verbose "calling Get-CMHealthCheck with parameter set"
        Get-CMHealthCheck @getParams
    }
    catch {
        Write-Error $_.Exception.Message
    }

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
                Healthcheckdebug = $Healthcheckfilename
                Healthcheckfilename = $Healthcheckfilename
                Verbose          = $VerbosePreference
            }
            Export-CMHealthCheck @expParams
        }
        else {
            throw "report folder not found: $ReportFolder"
        }
    }
    catch {
        Write-Error $_.Exception.Message
    }
}