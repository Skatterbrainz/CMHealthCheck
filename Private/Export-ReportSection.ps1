Function Export-ReportSection {
	param (
		[parameter()] $HealthCheckXML,
		[parameter()] $Section,
		[parameter()] $SqlConn,
		[parameter()][string] $SiteCode,
		[parameter()][int] $NumberOfDays,
		[parameter()][string] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()] $ReportTable,
		[parameter()][switch] $Detailed
	)
	Write-Log -Message "(Export-ReportSection): Section = $Section" -LogFile $logfile
	if ($Detailed) {
		Write-Log -Message "detailed...... True" -LogFile $logfile
		Write-Log -Message "section....... $Section" -LogFile $logfile
		Write-Log -Message "numberofdays.. $NumberOfDays" -LogFile $logfile
		Write-Log -Message "sitecode...... $SiteCode" -LogFile $logfile
	}
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.IsTextOnly.ToLower() -eq 'true') { continue }
		if ($healthCheck.IsActive.ToLower() -ne 'true') { continue }
		if ($healthCheck.Section.ToLower() -ne $Section) { continue }
		$sqlquery  = $healthCheck.SqlQuery
		$tablename = (Set-ReplaceString -Value $healthCheck.XMLFile -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $servername)
		$xmlTableName = $healthCheck.XMLFile
		if ($Section -eq 5) {
			if (!($Detailed)) {
				$tablename += "summary"
				$xmlTableName += "summary"
				$gbfiels = ""
				foreach ($field in $healthCheck.Fields.Field) {
					if ($field.groupby -in ("2")) {
						if ($gbfiels.Length -gt 0) { $gbfiels += "," }
						$gbfiels += $field.FieldName
					}
				}
				$sqlquery = "select $($gbfiels), count(1) as Total from ($($sqlquery)) tbl group by $($gbfiels)"
			} else {
				$tablename += "detail"
				$xmlTableName += "detail"
				$sqlquery = $sqlquery -replace "select distinct", "select"
				$sqlquery = $sqlquery -replace "select", "select distinct"
			}
		}
		$filename = Join-Path -Path $reportFolder -ChildPath ($tablename + '.xml')
		$row = $ReportTable.NewRow()
		$row.TableName = $xmlTableName
		$row.XMLFile = $tablename + ".xml"
		$ReportTable.Rows.Add($row)
		Write-Log -Message "XMLFile.... $filename" -LogFile $logfile
		Write-Log -Message "Table...... $TableName - Information...Starting" -LogFile $logfile
		Write-Log -Message "Type....... $($healthCheck.querytype)" -LogFile $logfile
		try {
			switch ($healthCheck.querytype.ToLower()) {
				'mpconnectivity' {
					Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mplist' | Out-Null}
				'mpcertconnectivity' {
					Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mpcert' | Out-Null}
				'sql' {
					Get-SQLData -sqlConn $sqlConn -SQLQuery $sqlquery -FileName $fileName -TableName $tablename -siteCode $siteCode -NumberOfDays $NumberOfDays -servername $servername -healthcheck $healthCheck -logfile $logfile -section $section -detailed $detailed | Out-Null}
				'baseosinfo' {
					Write-BaseOSInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null}
				'diskinfo' {
					Write-DiskInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null}
				'networkinfo' {
					Write-NetworkInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null}
				'rolesinstalled' {
					Write-RolesInstalled -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile | Out-Null}
				'servicestatus' {
					#Write-ServiceStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null}
					Write-Services -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null}
				'hotfixstatus' {
					if (-not $NoHotfix) {
						Write-HotfixStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -ContinueOnError $true | Out-Null
					}
				}
				'discoveries' {
					Write-DiscoveryMethods -FileName $filename -TableName $tablename -sitecode $SiteCode -sqlConn $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'devcollections' {
					Write-DevCollections -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'usercollections' {
					Write-UserCollections -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'packages' {
					Write-CmPackages -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'boundarygroups' {
					Write-BoundaryGroups -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'boundaries' {
					Write-Boundaries -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $SqlConn.datasource -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'localgroups' {
					Write-LocalGroups -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $servername -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'localusers' {
					Write-LocalUsers -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $servername -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'installedapps' {
					Write-InstalledApps -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $servername -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				'sqlmemory' {
					Write-SqlMemory -FileName $filename -TableName $tablename -sitecode $SiteCode -ServerName $servername -LogFile $logfile -ContinueOnError $True | Out-Null
				}
				default {}
			}
		} catch {
			$errorMessage = $Error[0].Exception.Message
			$errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred..." -Severity 3 -LogFile $logfile
			Write-Log -Message "Error $errorCode : $errorMessage connecting to $servername" -Severity 3 -LogFile $logfile
			$Error.Clear()
		}
		Write-Log -Message "$tablename Information...Done" -LogFile $logfile
	}
	Write-Log -Message "EndSection. $section ***" -LogFile $logfile
}
