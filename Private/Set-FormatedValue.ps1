Function Set-FormatedValue {
	param (
		[parameter()] $Value,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $Format,
		[parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $SiteCode
	)
	# Backward-compatible alias wrapper for misspelled helper name.
	Set-FormattedValue -Value $Value -Format $Format -SiteCode $SiteCode
}