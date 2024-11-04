$wdAlignPageNumberLeft   = 0
$wdAlignPageNumberCenter = 1
$wdAlignPageNumberRight  = 2

$wdPageNumberStyleArabic = 0
$wdPageNumberStyleUppercaseRoman = 1
$wdPageNumberStyleLowercaseRoman = 2
$wdPageNumberStyleUppercaseLetter = 3
$wdPageNumberStyleLowercaseLetter = 4

$wdAlignParagraphLeft    = 0
$wdAlignParagraphCenter  = 1
$wdAlignParagraphRight   = 2
$wdAlignParagraphJustify = 3

function Set-WordFormatting {
	Write-Log -Message "function... Set-WordFormatting ****" -LogFile $logfile
	if ($WordVersion -ge "16.0") {
		Write-Log -Message "setting styles for Word 2016" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	} elseif ($WordVersion -eq "15.0") {
		Write-Log -Message "setting styles for Word 2013" -LogFile $logfile
		$x1 = "Grid Table 4 - Accent 1"
		$x2 = "Grid Table 4 - Accent 1"
	} elseif ($WordVersion -eq "14.0") {
		Write-Log -Message "setting styles for Word 2010" -LogFile $logfile
		$x1 = "Medium Shading 1 - Accent 1"
		$x2 = "Light Grid - Accent 1"
	}
	Write-Output @($x1, $x2)
}
function Set-WordAbstract {
	$absText1 = "This document provides a point-in-time report of the current state of the "
	$absText1 += "System Center Configuration Manager site environment for $CustomerName. "
	$absText1 += "For questions, concerns or comments, please consult the author of this "
	$absText1 += "assessment report."
	$absText2 = "This report was generated using CMHealthCheck $ModuleVer on $(Get-Date)."

	Write-WordText -WordSelection $selection -Text "Abstract" -Style "Heading 1" -NewLine $true
	Write-WordText -WordSelection $selection -Text $absText1 -NewLine $true
	Write-WordText -WordSelection $selection -Text $absText2 -NewLine $true
}

function Set-WordOptions {
	Write-Log -Message "configuring word options for current session" -LogFile $logfile
	$Word.Options.CheckGrammarAsYouType  = $False
	$Word.Options.CheckSpellingAsYouType = $False
	$Doc.Styles("Normal").Font.Size = $NormalFontSize
}

function Set-WordTOC {
	Write-Log -Message "inserting table of contents" -LogFile $logfile
	$toc = $BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.Insert($selection.Range,$True) | Out-Null
}

function Set-WordFooter {
	Write-Log -Message "writing document footer content..."
	if ($Template -eq "") {
		$selection.HeaderFooter.Range.Text= "Copyright $([char]0x00A9) $((Get-Date).Year) - $CopyrightName"
	}
	$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null
}

function Write-WordText {
	param (
		[parameter()] $WordSelection,
		[parameter()][string] $Text = "",
		[parameter()][string] $Style = "Normal",
		[parameter()] $Bold    = $false,
		[parameter()] $NewLine = $false,
		[parameter()] $NewPage = $false
	)
	Write-Log -Message "(Write-WordText)" -LogFile $logfile
	$texttowrite = ""
	$wordselection.Style = $Style
	if ($bold) {
		$wordselection.Font.Bold = 1
	} else {
		$wordselection.Font.Bold = 0
	}
	$texttowrite += $text
	$wordselection.TypeText($text)
	If ($newline) { $wordselection.TypeParagraph() }
	If ($newpage) { $wordselection.InsertNewPage() }
}

function Set-WordDocumentProperty {
	param (
		[parameter(Mandatory)] $Document,
		[parameter(Mandatory)] [string] $Name,
		[parameter(Mandatory)] [string] $Value
	)
	Write-Log -Message "info: document property [$Name] set to [$Value]" -LogFile $logfile
	$document.BuiltInDocumentProperties($Name) = $Value
}

function Set-DocAppendix {
	param ()
	Write-Log -Message "(Set-DocAppendix)" -LogFile $logfile
	$appendix = @(
		("ConfigMgr Hardware Recommendations", "https://technet.microsoft.com/en-us/library/mt589500.aspx#bkmk_ScaleSieSystems"),
		("ConfigMgr Supported Operating Systems", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/supported-operating-systems-for-site-system-servers"),
		("ConfigMgr Supported SQL Server Versions", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/support-for-sql-server-versions"),
		("ConfigMgr Internet-Based Client Management", "https://docs.microsoft.com/en-us/sccm/core/clients/manage/plan-internet-based-client-management"),
		("ConfigMgr Site Size and Scale Information", "https://docs.microsoft.com/en-us/sccm/core/plan-design/configs/size-and-scale-numbers"),
		("ConfigMgr Support Lifecycle Information", "https://support.microsoft.com/en-us/lifecycle/search?alpha=Microsoft%20System%20Center%202012%20Configuration%20Manager"),
		("ConfigMgr Client Installation Properties", "https://docs.microsoft.com/en-us/sccm/core/clients/deploy/about-client-installation-properties"),
		("Best Practices for Managing Software Updates", "https://docs.microsoft.com/en-us/sccm/sum/plan-design/software-updates-best-practices"),
		("Deploy Windows 10 with MDT", "https://docs.microsoft.com/en-us/windows/deployment/deploy-windows-mdt/deploy-windows-10-with-the-microsoft-deployment-toolkit"),
		("Blogs - WindowsNoob", "https://www.windows-noob.com/forums/portal"),
		("Blogs - Deployment Research", "https://deploymentresearch.com/"),
		("Blogs - SC ConfigMgr", "https://www.scconfigmgr.com")
	)
	
	Write-Log -Message "inserting document Appendix..." -LogFile $logfile
	Write-WordText -WordSelection $selection -Text "Appendix A - Resource References" -Style "Heading 1" -NewLine $true
	$selection.TypeParagraph()

	foreach ($app in $appendix) {
		$caption = $app[0]
		$link = $app[1]
		Write-WordText -WordSelection $selection -Text $caption -NewLine $true
		Write-WordText -WordSelection $selection -Text $link -NewLine $true
		$selection.TypeParagraph()
	}
}

function Set-DocProperties {
	Write-Log -Message "(Set-DocProperties)" -LogFile $logfile
	if ($bAutoProps -eq $True) {
		Write-Log -Message "setting document properties" -LogFile $logfile
		$doc.BuiltInDocumentProperties("Title")    = "System Center Configuration Manager HealthCheck"
		$doc.BuiltInDocumentProperties("Subject")  = "Prepared for $CustomerName"
		$doc.BuiltInDocumentProperties("Author")   = $AuthorName
		$doc.BuiltInDocumentProperties("Company")  = $CopyrightName
		$doc.BuiltInDocumentProperties("Category") = "REPORTS"
		$doc.BuiltInDocumentProperties("Keywords") = "sccm,healthcheck,systemcenter,configmgr,$CustomerName"
	}
}

function Write-DocReportSections {
	Write-Log -Message "(Write-DocReportSections)" -LogFile $logfile 
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '1' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '2' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '3' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '4' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -section '5' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	
	if ($detailed -eq $true) {
		Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '5' -Detailed $true -Doc $doc -Selection $selection -LogFile $logfile 
		Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
	}
	
	Write-WordReportSection -HealthCheckXML $HealthCheckXML -Section '6' -Doc $doc -Selection $selection -LogFile $logfile 
	Write-WordTableGrid -Caption "Review Comments" -Rows 3 -ColumnHeadings $ReviewTableCols -StyleName $ReviewTableStyle
}

Function Write-WordReportSection {
	param (
		[parameter()] $HealthCheckXML,
		[parameter()] $Section,
		[parameter()] $Detailed = $false,
		[parameter()] $Doc,
		[parameter()] $Selection,
		[parameter()] $LogFile
	)
	Write-Log -Message "(Write-WordReportSection)" -LogFile $logfile
	Write-Log -Message "section....... $section" -LogFile $logfile
	Write-Log -Message "detail........ $($detailed.ToString())" -LogFile $logfile

	foreach ($healthCheck in $HealthCheckXML.dtsHealthCheck.HealthCheck) {
		if ($healthCheck.Section.ToLower() -ne $Section) { continue }
		$Description = $healthCheck.Description -replace("@@NumberOfDays@@", $NumberOfDays)
		if ($healthCheck.IsActive.ToLower() -ne 'true') { continue }
		if ($healthCheck.IsTextOnly.ToLower() -eq 'true') {
			if ($Section -eq 5) {
				if ($detailed -eq $false) {
					$Description += " - Overview"
				} else {
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
			} else {
				$tablename += "detail"
			}
		}
		foreach ($rp in $ReportTable) {
			if ($rp.TableName -eq $tableName) {
				$bFound = $true
				$filename = $rp.XMLFile
				Write-Log -Message "xmlfile....... $filename" -LogFile $logfile
				if ($filename.IndexOf("_") -gt 0) {
					$xmltitle = $filename.Substring(0,$filename.IndexOf("_"))
					$xmltile = ($rp.TableName.Substring(0,$rp.TableName.IndexOf("_")).Replace("@","")).Tolower()
					switch ($xmltile) {
						"sitecode"   { $xmltile = "Site Code: " }
						"servername" { $xmltile = "Server Name: " }
					}
					switch ($healthCheck.WordStyle) {
						"Heading 1" { $newstyle = "Heading 2" }
						"Heading 2" { $newstyle = "Heading 3" }
						"Heading 3" { $newstyle = "Heading 4" }
						default { $newstyle = $healthCheck.WordStyle }
					}
					$xmltile += $filename.Substring(0,$filename.IndexOf("_"))
					Write-WordText -WordSelection $selection -Text $xmltile -Style $newstyle -NewLine $true
				}
				if (!(Test-Path ($reportFolder + $filename))) {
					Write-WordText -WordSelection $selection -Text $healthCheck.EmptyText -NewLine $true
					Write-Log -Message "Table does not exist" -LogFile $logfile -Severity 2
					$selection.TypeParagraph()
				} else {
					#Write-Log -Message "importing XML file: $filename" -LogFile $logfile
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
							Write-Log -Message "table style... $TableStyle" -LogFile $logfile
							$table.Style = $TableStyle
							# added to force table width consistency in 1.0.4 (Issue 13)
							$table.PreferredWidthType = 2
							$table.PreferredWidth = 100
							$i = 1;
							Write-Log -Message "structure..... $count rows and $Columns columns" -LogFile $logfile
							Write-Log -Message "writing table column headings..." -LogFile $logfile
							foreach ($field in $HealthCheck.Fields.Field) {
								if ($section -eq 5) {
									if (($detailed) -and ($field.groupby -notin ('1','2'))) { continue }
									elseif ((!($detailed)) -and ($field.groupby -notin ('2','3'))) { continue }
								}
								$Table.Cell(1, $i).Range.Font.Bold = $True
								$Table.Cell(1, $i).Range.Text = $field.Description
								#Write-Log -Message "--column: $($field.Description)" -LogFile $logfile
								$i++
							} # foreach
							$xRow = 2
							$records = 1
							$y=0
							Write-Log -Message "writing data rows for table body..." -LogFile $logfile
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
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($row.$($field.FieldName))
										}
										default {
											$TextToWord = $row.$($field.FieldName);
										}
									}
									#Write-Log -Message "--value: $($TextToWord.ToString())" -LogFile $logfile
									if ([string]::IsNullOrEmpty($TextToWord)) {
										$TextToWord = " "
										$val = " "
									} elseif (Test-Numeric $TextToWord) {
										#Write-Log -Message "rounding numeric value precision" -LogFile $logfile
										$val = ([math]::Round($TextToWord,2)).ToString()
									} else {
										$val = $TextToWord.ToString()
									}
									$Table.Cell($xRow, $i).Range.Text = $val
									$i++
								} # foreach
								$xRow++
								$records++
							} # foreach
							$selection.EndOf(15) | Out-Null
							$selection.MoveDown() | Out-Null
							$doc.ActiveWindow.ActivePane.view.SeekView = 0
							$selection.EndKey(6, 0) | Out-Null
							if ($count -gt 2) {
								Write-Verbose "SORT OPERATION - SORTING TABLE"
								$Tables.Sort
								Write-Log -Message "NEW: appending row count label below table" -LogFile $logfile
								Write-WordText -WordSelection $selection -Text "$count items found" -Style "Normal" -NewLine $true
								$selection.TypeParagraph()
							}
							$selection.TypeParagraph()
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
							# added to force table width consistency in 1.0.4 (Issue 13)
							$table.PreferredWidthType = 2
							$table.PreferredWidth = 100
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
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.Rows[0].$($field.FieldName))
										}
										default {
											$TextToWord = $datatable.Rows[0].$($field.FieldName)
										}
									} # switch
									if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									$Table.Cell($i, 2).Range.Text = $TextToWord.ToString()
								} else {
									$TextToWord = "";
									switch ($field.Format.ToLower()) {
										"message" {
											$TextToWord = Get-MessageInformation -MessageID ($datatable.$($field.FieldName))
										}
										"messagesolution" {
											$TextToWord = Get-MessageSolution -MessageID ($datatable.$($field.FieldName))
										}
										default {
											$TextToWord = $datatable.$($field.FieldName)
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
										}
										"messagesolution" {
											$TextToWord = ($field.Description + " : " + (Get-MessageSolution -MessageID ($row.$($field.FieldName))))
										}
										default {
											$TextToWord = ($field.Description + " : " + $row.$($field.FieldName))
										}
									} # switch
									if ([string]::IsNullOrEmpty($TextToWord)) { $TextToWord = " " }
									Write-WordText -WordSelection $selection -Text ($TextToWord.ToString()) -NewLine $true
								} # foreach
								$selection.TypeParagraph()
								$records++
							} # foreach
						} # end of default switch case
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

function Write-WordTableGrid {
	param (
		[parameter(Mandatory)][string] $Caption,
		[parameter(Mandatory)][int] $Rows,
		[parameter(Mandatory)][string[]] $ColumnHeadings,
		[parameter()][string] $StyleName = $DefaultTableStyle
	)
	Write-Log -Message "(Write-WordTableGrid) Caption = $Caption" -LogFile $logfile
	$Selection.TypeText($Caption)
	if ($Caption -eq 'Review Comments') {
		$Selection.Style = "Heading 3"
	} else {
		$Selection.Style = "Heading 1"
	}
	$Selection.TypeParagraph()
	$Cols  = $ColumnHeadings.Length
	$Table = $doc.Tables.Add($Selection.Range, $rows, $cols)
	Write-Log -Message "table style: $StyleName" -LogFile $logfile
	$Table.Style = $StyleName
	for ($col = 1; $col -le $cols; $col++) {
		$Table.Cell(1, $col).Range.Text = $ColumnHeadings[$col-1]
	}
	for ($row = 1; $row -lt $rows; $row++) {
		$Table.Cell($row+1, 1).Range.Text = $row.ToString()
	}
	# set table width to 100%
	$Table.PreferredWidthType = 2
	$Table.PreferredWidth = 100
	# set column widths for more than 2 columns
	if ($cols -gt 2) {
		$Table.Columns.First.PreferredWidthType = 2
		if ($ColumnHeadings[0].length -lt 5) {
			# squeeze first column if heading is "No.", etc.
			$Table.Columns.First.PreferredWIdth = 5
		} else {
			$Table.Columns.First.PreferredWIdth = 7
		}
		$Table.Columns(2).PreferredWidthType = 2
		$Table.Columns(2).PreferredWIdth = 7
	} else {
		# set column widths for 1 or 2 columns only
		$Table.Columns.First.PreferredWidthType = 2
		$Table.Columns.First.PreferredWidth = 7
	}
	$Selection.EndOf(15) | Out-Null
	$Selection.MoveDown() | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = 0
	$Selection.EndKey(6, 0) | Out-Null
	$Selection.TypeParagraph()
}

function Get-WordTempSource {
	<#
	.SYNOPSIS
	Copy Source Document File to Destination
	
	.DESCRIPTION
	Copies a source DOCX file to a temporary name and returns the new filename
	
	.PARAMETER SourceFile
	Path and name of source document file
	
	.EXAMPLE
	$newfile = Get-WordTempSource -SourceFile "c:\files\myfile.docx"
	$newfile == "c:\users\johndoe\documents\cmhealthreport.docx"
	
	.NOTES
	#>
	param (
		[parameter(Mandatory=$True, HelpMessage="Name of Template File")]
		[ValidateNotNullOrEmpty()]
		[string] $SourceFile
	)
	if (Test-Path -Path $SourceFile) {
		$newFile = Join-Path -Path $OutputFolder -ChildPath $TempFilename
		Write-Log -Message "copying source [$Template] to temp file [$newFile]..." -LogFile $logfile
		try {
			Copy-Item -Path $Template -Destination $newFile -ErrorAction Stop
			$result = $True
		} catch {
			Write-Log -Message "ERROR: Failed to clone template from $Template" -Severity 3 -LogFile $logfile
			break
		}
	}
	Write-Output $newFile
}