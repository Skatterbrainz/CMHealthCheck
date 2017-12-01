function Write-WordTableGrid {
    param (
        [parameter(Mandatory=$True, HelpMessage="Table Caption Heading")]
            [string] $Caption,
        [parameter(Mandatory=$True, HelpMessage="Number of Rows")]
            [int] $Rows,
        [parameter(Mandatory=$True, HelpMessage="Array of Column Headings")]
            [string[]] $ColumnHeadings,
        [parameter(Mandatory=$False, HelpMessage="Table Style")]
            [string] $TableStyle = 'Grid Table 4 - Accent 1'
	)
	Write-Log -Message "inserting custom table: $Caption" -LogFile $logfile
    $Selection.TypeText($Caption)
    $Selection.Style = "Heading 1"
    $Selection.TypeParagraph()
    $Cols  = $ColumnHeadings.Length
    $Table = $doc.Tables.Add($Selection.Range, $rows, $cols)
    $Table.Style = $TableStyle
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
        $Table.Columns.First.PreferredWIdth = 7
        $Table.Columns(2).PreferredWidthType = 2
        $Table.Columns(2).PreferredWIdth = 7
    }
    # set column widths for 1 or 2 columns only
    else {
        $Table.Columns.First.PreferredWidthType = 2
        $Table.Columns.First.PreferredWidth = 7
    }
    $Selection.EndOf(15) | Out-Null
	$Selection.MoveDown() | Out-Null
	$doc.ActiveWindow.ActivePane.view.SeekView = 0
	$Selection.EndKey(6, 0) | Out-Null
	$Selection.TypeParagraph()
}