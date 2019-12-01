Function Set-FormatedValue {
    param (
        [parameter()] $Value,
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $Format,
        [parameter(Mandatory)][ValidateNotNullOrEmpty()][string] $SiteCode
	)
	Write-Log -Message "function... Set-FormatedValue ****" -LogFile $logfile
	Write-Log -Message "format..... $Format" -LogFile $logfile
	Write-Log -Message "sitecode... $SiteCode" -LogFile $logfile
	if ($null -eq $Value) {
		Write-Log -Message "value...... NULL" -LogFile $logfile
	}
	else {
		Write-Log -Message "value...... $Value" -LogFile $logfile
	}
	switch ($Format.ToLower()) {
		'schedule' {
			$schedule_Class = [wmiclass]""
			$schedule_class.psbase.path = "\\$($smsprovider)\root\sms\site_$($SiteCodeNamespace):SMS_ScheduleMethods"
			$schedule = ($schedule_class.ReadFromString($value)).TokenData
			if ($schedule.DaySpan -ne 0) { $return = ($schedule.DaySpan * 24 * 60) }
			elseif ($schedule.HourSpan -ne 0) { $return = ($schedule.HourSpan * 60) }
			elseif ($schedule.MinuteSpan -ne 0) { $return = ($schedule.MinuteSpan) }
			return $return
		}
        'alertsname' {
			if ($null -eq $value) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'$databasefreespacewarning' { $return = 'Low free space alert for database on site' }
					'$sumcompliance2updategroupdeploymentname' { $return = 'Low deployment success rate alert of update group' }
					default { $return = $value }
				}
			}
            return $return
        }
        'alertsseverity' {
			if ($null -eq $value) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'1' { $return = 'Error' }
					'2' { $return = 'Warning' }
					'3' { $return = 'Informational' }
					default { $return = 'Unknown' }
				}
			}
            return $return
        }
        'alertstypeid' {
            switch ($value.ToString().ToLower()) {
                '12' { $return = 'Update group deployment success' }
                '25' { $return = 'Database free space warning' }
                '31' { $return = 'Malware detection' }
                default { $return = $value }
            }
            Write-Output $return
        }
		'messagesolution' {
			Write-Log -Message "[messagesolution] convert to string" -LogFile $logfile
			if ($null -ne $value) {
				$return = $value.ToString()
			}
			Write-Output $return
		}
		default {
			Write-Output $value
		}
	}
}