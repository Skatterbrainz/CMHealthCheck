function Test-Numeric ($x) {
	($x -match '^\d+$')
	# try {
	# 	0 + $x | Out-Null
	# 	return $true
	# } catch {
	# 	return $false
	# }
}