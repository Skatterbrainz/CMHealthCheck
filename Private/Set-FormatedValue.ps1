Function Set-FormatedValue {
    param (
        $Value,
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
	        [string] $Format,
        [parameter(Mandatory=$True)]
            [ValidateNotNullOrEmpty()]
            [string] $SiteCode
	)
	Write-Log -Message "function... Set-FormatedValue ****" -LogFile $logfile
	Write-Log -Message "format..... $Format" -LogFile $logfile
	Write-Log -Message "sitecode... $SiteCode" -LogFile $logfile
	if ($Value -eq $null) {
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
			break
		}
        'alertsname' {
			if ($value -eq $null) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'$databasefreespacewarning' {
						$return = 'Low free space alert for database on site'
						break
					}
					'$sumcompliance2updategroupdeploymentname' {
						$return = 'Low deployment success rate alert of update group'
						break
					}
					default {
						$return = $value
						break
					}
				}
			}
            return $return
            break
        }
        'alertsseverity' {
			if ($value -eq $null) {
				$return = ''
			}
			else {
				switch ($value.ToString().ToLower()) {
					'1' {
						$return = 'Error'
						break
					}
					'2' {
						$return = 'Warning'
						break
					}
					'3' {
						$return = 'Informational'
						break
					}
					default {
						$return = 'Unknown'
						break
					}
				}
			}
            return $return
            break
        }
        'alertstypeid' {
            switch ($value.ToString().ToLower()) {
                '12' {
                    $return = 'Update group deployment success'
                    break
                }
                '25' {
                    $return = 'Database free space warning'
                    break
                }
                '31' {
                    $return = 'Malware detection'
                    break
                }
                default {
                    $return = $value
                    break
                }
            }
            Write-Output $return
            break
        }
		'messagesolution' {
			Write-Log -Message "[messagesolution] convert to string" -LogFile $logfile
			if ($value -ne $null) {
				$return = $value.ToString()
			}
			Write-Output $return
			break
		}
		default {
			Write-Output $value
			break
		}
	}
}