Function Write-WordReportSection {
    param (
		$HealthCheckXML,
		$Section,
		[parameter(Mandatory=$False)]
		$Detailed = $false,
        $Doc,
		$Selection,
        $LogFile
	)
	Write-Log -Message "function...... Write-WordReportSection ****" -LogFile $logfile
	Write-Log -Message "section....... $section" -LogFile $logfile
	Write-Log -Message "detail........ $($detailed.ToString())" -LogFile $logfile
	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.Section.tolower() -ne $Section) { continue }
		$Description = $healthCheck.Description -replace("@@NumberOfDays@@", $NumberOfDays)
		if ($healthCheck.IsActive.tolower() -ne 'true') { continue }
        if ($healthCheck.IsTextOnly.tolower() -eq 'true') {
            if ($Section -eq 5) {
                if ($detailed -eq $false) { 
                    $Description += " - Overview" 
                } 
                else { 
                    $Description += " - Detailed"
                }            
            }
			Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
			Continue;
		}
		Write-Log -Message "description... $Description" -LogFile $logfile
		Write-WordText -WordSelection $selection -Text $Description -Style $healthCheck.WordStyle -NewLine $true
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
				$bFound = $true
				Write-Log -Message "xmlfile....... $($rp.XMLFile)" -LogFile $logfile
				$filename = $rp.XMLFile				
				if ($filename.IndexOf("_") -gt 0) {
					$xmltitle = $filename.Substring(0,$filename.IndexOf("_"))
					$xmltile = ($rp.TableName.Substring(0,$rp.TableName.IndexOf("_")).Replace("@","")).Tolower()
					switch ($xmltile) {
						"sitecode"   { $xmltile = "Site Code: "; break; }
						"servername" { $xmltile = "Server Name: "; break; }
					}
					switch ($healthCheck.WordStyle) {
						"Heading 1" { $newstyle = "Heading 2"; break; }
						"Heading 2" { $newstyle = "Heading 3"; break; }
						"Heading 3" { $newstyle = "Heading 4"; break; }
						default { $newstyle = $healthCheck.WordStyle; break }
					}
					$xmltile += $filename.Substring(0,$filename.IndexOf("_"))
					Write-WordText -WordSelection $selection -Text $xmltile -Style $newstyle -NewLine $true
				}
				
	            if (!(Test-Path ($reportFolder + $rp.XMLFile))) {
					Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
					Write-Log -Message "Table does not exist" -LogFile $logfile -Severity 2
					$selection.TypeParagraph()
				}
				else {
					Write-Log -Message "importing XML file: $filename" -LogFile $logfile
					$datatable = Import-CliXml -Path ($reportFolder + $filename)
					$count = 0
					$datatable | Where-Object { $count++ }
					
		            if ($count -eq 0) {
						Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
						Write-Log -Message "Table......... 0 rows" -LogFile $logfile -Severity 2
						$selection.TypeParagraph()
						continue
		            }

					switch ($healthCheck.PrintType.ToLower()) {
						"table" {
							Write-Log -Message "table type.... table" -LogFile $logfile
							$Table = $Null
					        $TableRange = $Null
					        $TableRange = $doc.Application.Selection.Range
                            $Columns = 0
                            foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            } # foreach
							$Table = $doc.Tables.Add($TableRange, $count+1, $Columns)
							$table.Style = $TableStyle
							$i = 1;
							Write-Log -Message "structure..... $count rows and $Columns columns" -LogFile $logfile
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }

								$Table.Cell(1, $i).Range.Font.Bold = $True
								$Table.Cell(1, $i).Range.Text = $field.Description
								$i++
	                        } # foreach
							$xRow = 2
							$records = 1
							$y=0
							foreach ($row in $datatable) {
								if ($records -ge 500) {
									Write-Log -Message ("Exported..... $(500*($y+1)) records") -LogFile $logfile
									$records = 1
									$y++
								}
								$i = 1;
								foreach ($field in $HealthCheck.Fields.Field) {
                                    if ($section -eq 5) {
                                        if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                        elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                    }
									$Table.Cell($xRow, $i).Range.Font.Bold = $false
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
                                    if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($xRow, $i).Range.Text = $TextToWord.ToString()
									$i++
		                        } # foreach
								$xRow++
								$records++
							} # foreach
							$selection.EndOf(15) | Out-Null
							$selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.view.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
						}
						"simpletable" {
							Write-Log -Message "table type.... simpletable" -LogFile $logfile
							$Table = $Null
							$TableRange = $Null
							$TableRange = $doc.Application.Selection.Range
							$Columns = 0
							foreach ($field in $HealthCheck.Fields.Field) {
                                if ($section -eq 5) {
                                    if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
                                    elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
                                }
                                $Columns++
                            } # foreach
							$Table = $doc.Tables.Add($TableRange, $Columns, 2)
							$table.Style = $TableSimpleStyle
							$i = 1;
							Write-Log -Message "structure..... $Columns rows and 2 columns" -LogFile $logfile
							$records = 1
							$y=0
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
								$Table.Cell($i, 1).Range.Font.Bold = $true
								$Table.Cell($i, 1).Range.Text = $field.Description
								$Table.Cell($i, 2).Range.Font.Bold = $false
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
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
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
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								}
								$i++
								$records++
							} # foreach
					        $selection.EndOf(15) | Out-Null
					        $selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.View.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							$selection.TypeParagraph()
							break
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
									Write-WordText -WordSelection $selection -Text ($TextToWord.ToString()) -NewLine $true
		                        } # foreach
								$selection.TypeParagraph()
								$records++
		                    } # foreach
						}
					} # switch
					Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
				}
			}
		} # foreach
        if ($bFound -eq $false) {
		    Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
		    Write-Log -Message ("Table does not exist") -LogFile $logfile -Severity 2
		    $selection.TypeParagraph()
		}
	} # foreach
}