Function Get-SQLData {
    param (
	    [parameter(Mandatory)] $sqlConn,
	    [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $SQLQuery,
	    [parameter()][string] $FileName,
	    [parameter()][string] $TableName,
	    [parameter()][string] $SiteCode,
	    [parameter()][int] $NumberOfDays,
	    [parameter()] $LogFile,
		[parameter()][string] $ServerName,
		[parameter()][bool] $ContinueOnError = $true,
		[parameter()] $HealthCheck,
        [parameter()] $Section,
        [parameter()][switch] $Detailed
	)
	Write-Log -Message "[function: Get-SQLData]"
	if ($Detailed) { 
        Write-Log -Message "  [detailed = True]" -LogFile $logfile
    }
    try {
        $SqlCommand = $sqlConn.CreateCommand()
		$logQuery       = Set-ReplaceString -value $SQLQuery -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $ServerName
		$executionquery = Set-ReplaceString -value $SQLQuery -SiteCode $SiteCode -NumberOfDays $NumberOfDays -ServerName $ServerName -space $false
        Write-Log -Message "SQL Query...`n$executionquery" -LogFile $logfile
	    Write-Log -Message "Log Query...`n$logQuery" -LogFile $logfile
        $SqlCommand.CommandTimeOut = 0
        $SqlCommand.CommandText = $executionquery
        $DataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $SqlCommand
        $dataset     = New-Object System.Data.Dataset
        $DataAdapter.Fill($dataset)
		if (($dataset.Tables.Count -eq 0) -or ($dataset.Tables[0].Rows.Count -eq 0)) { 
			Write-Log -Message "SQL Query returned 0 records" -LogFile $logfile
			Write-Log -Message "Table $tablename is empty. No file output to $filename ..." -LogFile $logfile
		}
		else {
			Write-Log -Message "SQL Query returned $($dataset.Tables[0].Rows.Count) records"
			foreach ($field in $healthCheck.Fields.Field) {
				Write-Log -Message ("   field = $($Field.FieldName) description = $($Field.Description)") -LogFile $logfile
                if ($section -eq 5) {
                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                    elseif (($detailed -eq $false) -and ($field.groupby -notin ('2','3'))) { continue }
                }
				if (![string]::IsNullOrEmpty($field.format)) {
					Write-Log -Message "   custom format specified for this attribute: $($Field.Format)" -LogFile $logfile
					foreach ($row in $dataset.Tables[0].Rows) {
						$tempx = Set-FormattedValue -Value $row.$($field.FieldName) -Format $field.format -SiteCode $SiteCode
						try {
							$row.$($field.FieldName) = $tempx
						}
						catch {
							$row
							break
						}
					}
				}
			}
			Write-Log -Message "Export: Exporting xml data to $filename" -LogFile $logfile
        	, $dataset.Tables[0] | Export-CliXml -Path $filename
		}
    }
    catch {
        $errorMessage = $Error[0].Exception.Message
        $errorCode = "0x{0:X}" -f $Error[0].Exception.ErrorCode
        if ($continueonerror -eq $false) { 
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred (stop)." -Severity 3 -LogFile $logfile
		}
        else { 
			Write-Log -Message "ERROR/EXCEPTION: The following error occurred (continue)." -Severity 3 -LogFile $logfile
		}
        Write-Log -Message "Error $errorCode : $errorMessage connecting to $ServerName" -Severity 3 -LogFile $logfile
	    $Error.Clear()
		Write-Log -Message "Unable to update file: $filename" -Severity 2 -LogFile $logfile
        if ($continueonerror -eq $false) {
			Throw "Error $errorCode : $errorMessage connecting to $ServerName"
		}
	}
}
