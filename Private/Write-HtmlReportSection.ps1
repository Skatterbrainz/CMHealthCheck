Function Write-HtmlReportSection {
    param (
		$HealthCheckXML,
		$Section,
		[switch] $Detailed,
        $LogFile
	)
    Write-Log -Message "---------------------------------------------------" -LogFile $logfile
	Write-Log -Message "function...... Write-HtmlReportSection ****" -LogFile $logfile
	Write-Log -Message "section....... $section" -LogFile $logfile
	Write-Log -Message "detail........ $($detailed.ToString())" -LogFile $logfile
	$result = ""
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.Section.ToLower() -ne $Section) { continue }
		$Description = $healthCheck.Description -replace("@@NumberOfDays@@", $NumberOfDays)
		if ($healthCheck.IsActive.ToLower() -ne 'true') { continue }
        if ($healthCheck.IsTextOnly.ToLower() -eq 'true') {
            if ($Section -eq 5) {
                if ($detailed -eq $false) { 
                    $Description += " - Overview" 
                } 
                else { 
                    $Description += " - Detailed"
                }            
            }
            $result += "<h3>$Description</h3>"
			#Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
			Continue;
        }
        Write-Log -Message "..................................................." -LogFile $logfile
		Write-Log -Message "description... $Description" -LogFile $logfile
        $result += "<h2>$Description</h2>"
        $bFound = $false
        $tableName = $healthCheck.XMLFile
        if ($Section -eq 5) {
            if (!($detailed)) { 
                $tablename += "summary" 
            } 
            else { 
                $tablename += "detail"
            }            
        }
		foreach ($rp in $ReportTable) {
			if ($rp.TableName -eq $tableName) {
				$bFound   = $true
				$filename = $rp.XMLFile
				Write-Log -Message "xmlfile....... $filename" -LogFile $logfile
				if ($filename.IndexOf("_") -gt 0) {
					$xmltitle = $filename.Substring(0,$filename.IndexOf("_"))
					$xmltile  = ($rp.TableName.Substring(0,$rp.TableName.IndexOf("_")).Replace("@","")).Tolower()
					switch ($xmltile) {
						"sitecode"   { $xmltile = "Site Code: "; break; }
						"servername" { $xmltile = "Server Name: "; break; }
					}
					switch ($healthCheck.WordStyle) {
						"Heading 1" { $CapStyle = "h2"; break; }
						"Heading 2" { $CapStyle = "h3"; break; }
						"Heading 3" { $CapStyle = "h4"; break; }
						default { $newstyle = $healthCheck.WordStyle; break }
					}
					$xmltile += $filename.Substring(0,$filename.IndexOf("_"))
					Write-Log -Message "--- xmlTile = $xmlTile" -LogFile $logfile
                    $result += "<$CapStyle>$xmlTile</$CapStyle>"
				}
				
	            if (!(Test-Path ($reportFolder + $filename))) {
                    $result += "<table class=`"reportTable`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
					Write-Log -Message "Table does not exist" -LogFile $logfile
				}
				else {
					Write-Log -Message "importing XML file: $filename" -LogFile $logfile
					$datatable = Import-CliXml -Path ($reportFolder + $filename)
					$count = 0
					$datatable | Where-Object { $count++ }
					
		            if ($count -eq 0) {
                        $result += "<table class=`"reportTable`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
						Write-Log -Message "Table......... 0 rows" -LogFile $logfile
						continue
		            }

					switch ($healthCheck.PrintType.ToLower()) {
						"table" {
							Write-Log -Message "table type.... table" -LogFile $logfile
                            $Columns = 0
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            } # foreach
							$i = 1;
							Write-Log -Message "--- structure..... $count rows and $Columns columns" -LogFile $logfile
                            Write-Log -Message "--- writing table column headings..." -LogFile $logfile
                            $table = "<table class=`"reportTable`"><tr>"
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $table += "<th class=`"headingRow1`">$($field.Description)</th>"
								$i++
                            } # foreach
                            $table += "</tr>"
							$records = 1
                            $rownum = 0
							$xRow = 2
                            $y = 0
							Write-Log -Message "--- writing data rows for table body..." -LogFile $logfile
							foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported..... $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								switch ($Script:TableRowStyle) {
									'Solid' {
										$table += "<tr class=`"rowstyle3`">"
										break
									}
									'Alternating' {
										if ($rownum % 2 -eq 0) {
											$table += "<tr class=`"rowstyle3`">"
										}
										else {
											$table += "<tr class=`"rowstyle4`">"
										}
										break
									}
									'Dynamic' {
										$table += "<tr class=`"rowstylex`">"
										break
									}
								}
								$i = 1;
								foreach ($field in $HealthCheck.Fields.Field) {
                                    if ($section -eq 5) {
                                        if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                        elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                    }
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($row.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($row.$($field.FieldName))
											break ;
										}										
										default {
											$TextToWord = $row.$($field.FieldName);
											break;
										}
									}
									if ([string]::IsNullOrEmpty($TextToWord)) { 
										$TextToWord = " " 
										$val = " "
									}
									elseif (Test-Numeric $TextToWord) {
										$val = ([math]::Round($TextToWord,2)).ToString()
									}
									else {
										$val = $TextToWord.ToString()
                                    }
                                    $table += "<td>$val</td>"
									$i++
		                        } # foreach
                                $table += "</tr>"
                                $records++
                                $rownum++
								$xRow++
							} # foreach
							Write-Log -Message "--- appending table row count: $Count" -LogFile $logfile
							$table += "<tr class=`"headingRow1`"><td colspan=`"$Columns`">$Count items found</td></tr></table>"
							$result += $table
                            <#
                            if ($count -gt 2) {
								Write-Verbose "SORT OPERATION - SORTING TABLE"
								#$Tables.Sort
								Write-Log -Message "NEW: appending row count label below table" -LogFile $logfile
                                #Write-WordText -WordSelection $selection -Text "$count items found" -Style "Normal" -NewLine $true
                                $result += "<p>$count items found</p>"
								#$selection.TypeParagraph()
                            }
                            #>
							break
						}
						"simpletable" {
							Write-Log -Message "table type.... simpletable" -LogFile $logfile
							$Columns = 0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }
							Write-Log -Message "structure..... $Columns rows and 2 columns" -LogFile $logfile
							$records = 1
							$rownum = 0
							$i = 1;
                            $y=0
                            $table = "<$CapStyle>$Caption</$CapStyle> <table class=`"reportTable`">"
							Write-Log -Message "--- building simpletable column heading cells" -Logfile $logfile
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
								if ($records -ge 500) {
									Write-Log -Message ("Exported..... $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
                                }
								if ($poshversion -ne 3) { 
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.Rows[0].$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.Rows[0].$($field.FieldName)
											break;
										}
									} # switch
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
								}
								else {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.$($field.FieldName))
											break ;
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.$($field.FieldName))
											break ;
										}											
										default {
											$TextToWord = $datatable.$($field.FieldName) 
											break;
										}
									} # switch
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
								}
								switch ($Script:TableRowStyle) {
									'Solid' {
										$table += "<tr class=`"rowstyle3`">"
										break
									}
									'Alternating' {
										if ($rownum % 2 -eq 0) {
											$table += "<tr class=`"rowstyle3`">"
										}
										else {
											$table += "<tr class=`"rowstyle4`">"
										}
										break
									}
									'Dynamic' {
										$table += "<tr class=`"rowstylex`">"
										break
									}
								}
                                $table += "<td class=`"rowstyle1`" style=`"width:300px`">$($field.Description)</td>"
                                $table += "<td>$($TextToWord.ToString())</td></tr>"
								$i++
                                $records++
                                $rownum++
							} # foreach
							Write-Log -Message "--- appending simpletable row count: $Count" -LogFile $logfile
							$table += "<tr class=`"headingRow1`"><td colspan=2>$Count items found</td></tr></table>"
							$result += $table
							break
						}
						default {
							Write-Log -Message "table type.... default" -LogFile $logfile
							$records = 1
							$y=0
		                    foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported...... $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
		                        foreach ($field in $HealthCheck.Fields.Field) {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = ($field.Description + " : " + (Get-MessageInformation -MessageID ($row.$($field.FieldName))))
											break ;
										}
										"messagesolution" {
											$TextToWord = ($field.Description + " : " + (Get-MessageSolution -MessageID ($row.$($field.FieldName))))
											break ;
										}												
										default {
											$TextToWord = ($field.Description + " : " + $row.$($field.FieldName))
											break;
										}
									} # switch
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
                                    $result += "<table class=`"reportTable`"><tr><td>$($TextToWord.ToString())</td></tr></table>"
		                        } # foreach
								$records++
		                    } # foreach
							#Write-Verbose "NEW: appending row count label below table"
                            $result += "<p>$($count+1) items found . . .</p>"
						} # end of default switch case
					} # switch
				}
			}
		} # foreach
        if ($bFound -eq $false) {
            $result += "<table class=`"reportTable`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
		    Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
		}
    } # foreach
    Write-Output $result
}
