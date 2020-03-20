function Write-CmPackages {
	param (
		[parameter(Mandatory)][string] $FileName,
		[parameter(Mandatory)][string] $TableName,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true
	)
	Write-Log -Message "(Write-CmPackages)" -LogFile $logfile
	$query = "select distinct PackageID, Name, 
	Case
		When (PackageType = 0)   Then 'Software Distribution Package'
		When (PackageType = 3)   Then 'Driver Package'
		When (PackageType = 4)   Then 'Task Sequence Package'
		When (PackageType = 5)   Then 'Software Update Package'
		When (PackageType = 6)   Then 'Device Settings Package'
		When (PackageType = 7)   Then 'Virtual Package'
		When (PackageType = 8)   Then 'Application'
		When (PackageType = 257) Then 'OS Image Package'
		When (PackageType = 258) Then 'Boot Image Package'
		When (PackageType = 259) Then 'OS Upgrade Package'
		WHEN (PackageType = 260) Then 'VHD Package'
		End as PkgType,
	PackageType, Description, SourceVersion as Version from dbo.v_Package order by Name"
	$packages = @(Invoke-DbaQuery -SqlInstance $ServerName -Database $SQLDBName -Query $query -ErrorAction SilentlyContinue)
	if ($null -eq $packages) { return }
	$Fields = @("Name","PkgID","Type","Description","Version")
	$pkgDetails = New-CmDataTable -TableName $tableName -Fields $Fields
	foreach ($pkg in $packages) {
		$row             = $pkgDetails.NewRow()
		$row.Name        = $pkg.Name
		$row.PkgID       = $pkg.PackageID
		$row.Type        = $pkg.PkgType
		$row.Version     = $pkg.Version
		$row.Description = $pkg.Description
		$pkgDetails.Rows.Add($row)
	}
	, $pkgDetails | Export-CliXml -Path ($filename)
}