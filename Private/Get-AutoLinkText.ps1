function Get-AutoLinkText {
    param (
        [parameter(Mandatory=$True)]
        [ValidateNotNullOrEmpty()]
        [string] $StringValue,
        [switch] $NewPage
    )
    $temp = @()
    $tokens = $StringValue.split(' ')
    $tokens | % { 
        if ($_.StartsWith('http')) {
            if ($NewPage) {
                $temp += "<a href=`"$_`" target=`"blank`">$_</a>"
            }
            else {
                $temp += "<a href=`"$_`">$_</a>"
            }
        }
        else {
            $temp += $_
        }
    }
    $temp -join ' '
}