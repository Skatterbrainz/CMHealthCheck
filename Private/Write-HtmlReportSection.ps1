Function Write-HtmlReportSection {
    param (
		$HealthCheckXML,
		$Section,
		[switch] $Detailed,
        $LogFile
    )
    Write-Log -Message "---------------------------------------------" -LogFile $logfile
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
            $result += "<p>$Description</p>"
			#Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
			Continue;
        }
        Write-Log -Message "..................................................." -LogFile $logfile
		Write-Log -Message "description... $Description" -LogFile $logfile
        #Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
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
                    #Write-WordText -WordSelection $selection -Text $xmltile -Style $newstyle -NewLine $true
                    $result += "<$CapStyle>$xmlTile</$CapStyle>"
				}
				
	            if (!(Test-Path ($reportFolder + $filename))) {
                    #Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
                    $result += "<table class=`"table1000`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
					Write-Log -Message "Table does not exist" -LogFile $logfile -Severity 2
				}
				else {
					#Write-Log -Message "importing XML file: $filename" -LogFile $logfile
					$datatable = Import-CliXml -Path ($reportFolder + $filename)
					$count = 0
					$datatable | Where-Object { $count++ }
					
		            if ($count -eq 0) {
                        $result += "<table class=`"table1000`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
						Write-Log -Message "Table......... 0 rows" -LogFile $logfile -Severity 2
						continue
		            }

					switch ($healthCheck.PrintType.ToLower()) {
						"table" {
                            Write-Log -Message "- - - - - - - - - - - - - - - - - -" -LogFile $logfile
							Write-Log -Message "table type.... table" -LogFile $logfile
                            $Columns = 0
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            } # foreach
                            Write-Verbose "[table] columns: $Columns"
                            $table = "<table class=`"table1000`">"
                            Write-Log -Message "table style... $TableStyle" -LogFile $logfile
							# added to force table width consistency in 1.0.4 (Issue 13)
							$i = 1;
							Write-Log -Message "structure..... $count rows and $Columns columns" -LogFile $logfile
                            Write-Log -Message "writing table column headings..." -LogFile $logfile
                            $table += "<tr>"
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $table += "<th class=`"columnstyle3`">$($field.Description)</th>"
								$i++
                            } # foreach
                            $table += "</tr>"
							$xRow = 2
							$records = 1
                            $y = 0
                            $rownum = 0
							Write-Log -Message "writing data rows for table body..." -LogFile $logfile
							foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported..... $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
                                }
                                if ($DynamicTableRows) {
                                    # alternate row colors even/odd
                                    if ($rownum % 2 -eq 0) {
                                        $table += "<tr class=`"rowstyle3`">"
                                    }
                                    else {
                                        $table += "<tr class=`"rowstyle4`">"
                                    }
                                }
                                else {
                                    $table += "<tr class=`"rowstylex`">"
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
								$xRow++
                                $records++
                                $rownum++
                                $table += "</tr>"
							} # foreach
                            $table += "</table>"
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
                            Write-Log -Message "- - - - - - - - - - - - - - - - - -" -LogFile $logfile
							Write-Log -Message "table type.... simpletable" -LogFile $logfile
							$Columns = 0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            }
                            $table = "<$CapStyle>$Caption</$CapStyle> <table class=`"table1000`"><tr>"
							$i = 1;
							Write-Log -Message "structure..... $Columns rows and 2 columns" -LogFile $logfile
							$records = 1
                            $y=0
                            $rownum = 0
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
                                $table += "<th class=`"$columnstyle3`">$($field.Description)</th>"
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
                                if ($DynamicTableRows) {
                                    if ($rownum % 2 -eq 0) {
                                        $table += "<tr class=`"rowstyle3`">"
                                    }
                                    else {
                                        $table += "<tr class=`"rowstyle4`">"
                                    }
                                }
                                else {
                                    $table += "<tr class=`"rowstylex`">"
                                }
                                $table += "<td style=`"width:200px`">$($field.Description)</td>"
                                $table += "<td>$($TextToWord.ToString())</td></tr>"
								$i++
                                $records++
                                $rownum++
                            } # foreach
                            $result += "<tr class=`"rowstyle4`"><td colspan=`"2`">$count items found</td></tr></table>"
							break
						}
						default {
                            Write-Log -Message "- - - - - - - - - - - - - - - - - -" -LogFile $logfile
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
                                    #Write-WordText -WordSelection $selection -Text ($TextToWord.ToString()) -NewLine $true
                                    $result += "<p>$($TextToWord.ToString())</p>"
		                        } # foreach
								$records++
		                    } # foreach
							#Write-Verbose "NEW: appending row count label below table"
                            $result += "<p>$($count+1) items found</p>"
						} # end of default switch case
					} # switch
                    #Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
                    $result += New-HtmlTableBlock -Caption "Review Comments" -CaptionStyle "h3" -TableStyle "width:1000px" -HeadingStyle "background:#c0c0c0" -HeadingNames "Item=60,Severity=100,Description" -RowStyle2 "background:#eee" -Rows 3
				}
			}
		} # foreach
        if ($bFound -eq $false) {
            $result += "<table style=`"width:800px`"><tr><td>$($healthCheck.EmptyText)</td></tr></table>"
		    Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
		}
    } # foreach
    Write-Output $result
}
