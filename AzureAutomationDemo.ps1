#requires -RunAsAdministrator
[CmdletBinding()]
param (
	[parameter()][string] $SiteServer = "cm01.contoso.local",
	[parameter()][string] $SiteCode = "P01",
	[parameter()][string] $Customer = "Contoso"
)

#$ErrorActionPreference = 'stop'
$VerbosePreference = 'silentlycontinue'

$SendFrom = Get-AutomationVariable -Name 'MailSender'
$SendTo   = Get-AutomationVariable -Name 'MailRecipients'
$Subject  = "CMHealthCheck Report $(Get-Date -f MM-dd-yyyy)"
$MsgBody  = "$Subject`nSite Server: $env:COMPUTERNAME"

if ([string]::IsNullOrEmpty($SendFrom)) {
	Write-Output "Automation Variable [SendFrom] is not defined in this Automation Account"
	exit
}
if ([string]::IsNullOrEmpty($SendTo)) {
	Write-Output "Automation Variable [SendTo] is not defined in this Automation Account"
	exit
}

function Send-Email {
	param (
		[parameter()] $Attachments
	)
	$sgCred   = Get-AutomationPSCredential -Name 'SendGridAccount'
	$userName = $sgCred.UserName
	$securePassword = $sgCred.Password
	$password   = $sgCred.GetNetworkCredential().Password
	$pwd = ConvertTo-SecureString $Password -AsPlainText -Force
	$Credential = New-Object System.Management.Automation.PSCredential $userName, $pwd

	Write-Output "Sending an email to $SendTo :: subject = $Subject"
	Send-MailMessage -SmtpServer "smtp.sendgrid.net" -Credential $Credential -UseSSL -Port 587 `
		-From $SendFrom -To $SendTo -Subject $Subject -Body $MsgBody `
		-Attachments $Attachments
}

function Install-PSModule {
	if (Get-Module 'CMHealthCheck' -ListAvailable) {
		Write-Verbose "** module is already installed"
		Write-Output 0
	}
	else {
		try {
			Write-Verbose "** installing module: CMHealthCheck"
			Install-Module 'CMHealthCheck' -AllowClobber
			Write-Output 0
		}
		catch {
			Write-Verbose "** Error: $($_.Exception.Message)"
			Write-Output 1
		}
	}
}

if ($(Install-PSModule) -eq 0) {
	#Write-Output "** setting execution policy to ByPass"
	#Set-ExecutionPolicy ByPass -Force
	Import-Module CMHealthCheck -Force
	$VerbosePreference = 'continue'
	Write-Output "** module was installed successfully. Running audit and report..."
	Write-Output (Get-Module "CMHealthCheck" -ListAvailable | Sort-Object Version -Descending | Select -First 1).Version -join '.'
	Write-Output "** module version is $mver"
	$VerbosePreference = 'silentlycontinue'
	$DataFolder = "c:\windows\temp"
	$ReportFolder = Join-Path $DataFolder "$(Get-Date -f 'yyyy-MM-dd')\$SmsProvider"
	Write-Output "** report folder path = $ReportFolder"

	Invoke-CMHealthCheck -SmsProvider $SiteServer `
		-CustomerName $Customer `
		-AuthorName "Skatterbrainz" `
		-CopyrightName "Skatterbrainz LPC" `
		-DataFolder $env:TEMP `
		-PublishFolder $env:TEMP `
		-Overwrite `
		-NoHotfix `
		-Detailed
	Write-Output "** data collection and export completed"
	$reportFile = Join-Path -Path "c:\windows\temp" -ChildPath $("cmhealthreport`-$SiteServer-$(Get-Date -f 'yyyyMMdd').htm")
	# example: "cmhealthreport-cm01.contoso.local-20191201.htm"
	Write-Output "** checking for report file: $reportFile"
	if (Test-Path $reportFile) {
		Send-Email -Attachments $reportFile
	}
	else {
		Write-Output "error - report file not found: $reportFile"
	}
}
else {
	Write-Output "** oooooh... massive toilet bowl implosion explosion!"
}