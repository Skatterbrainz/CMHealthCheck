Function ReportSection {
    param (
	    $HealthCheckXML,
		$Section,
		$SqlConn,
		[string] $SiteCode,
		$NumberOfDays,
		[string] $LogFile,
		[string] $ServerName,
		$ReportTable,
		[switch] $Detailed
	)
	Write-Log -Message "function... ReportSection ****" -LogFile $logfile
	if ($Detailed) { 
		Write-Log -Message "detailed... True" -LogFile $logfile
		Write-Log -Message "section.... $Section" -LogFile $logfile
		Write-Log -Message "numberofdays $NumberOfDays" -LogFile $logfile
		Write-Log -Message "sitecode... $SiteCode" -LogFile $logfile
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
            } 
            else { 
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
				'mpconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mplist' | Out-Null}
				'mpcertconnectivity' { Write-MPConnectivity -FileName $filename -TableName $tablename -sitecode $SiteCode -SiteCodeQuery $SiteCodeQuery -NumberOfDays $NumberOfDays -logfile $logfile -type 'mpcert' | Out-Null}
				'sql' { Get-SQLData -sqlConn $sqlConn -SQLQuery $sqlquery -FileName $fileName -TableName $tablename -siteCode $siteCode -NumberOfDays $NumberOfDays -servername $servername -healthcheck $healthCheck -logfile $logfile -section $section -detailed $detailed | Out-Null}
				'baseosinfo' { Write-BaseOSInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'diskinfo' { Write-DiskInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'networkinfo' { Write-NetworkInfo -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'rolesinstalled' { Write-RolesInstalled -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile | Out-Null}
				'servicestatus' { Write-ServiceStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null}
				'hotfixstatus' { 
                    if (-not $NoHotfix) {
                        Write-HotfixStatus -FileName $filename -TableName $tablename -sitecode $SiteCode -NumberOfDays $NumberOfDays -servername $servername -logfile $logfile -continueonerror $true | Out-Null
                    }
                }
           		default {}
			}
		}
		catch {
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
