'Variables available only for this group can be declared here.

Dim TargetFile, DBConnectionName, TableName, DateFrom, DateTo, AvgGroupPeriod, tagProgress, strSelectColumns

'The code configured here is executed while the condition configured in the Execution field is TRUE.

$RunScriptExport = 0

TargetFile = $TargetFileName

DBConnectionName = "DB"

Select Case $AreaExportar
	Case 1 TableName = "Area100Instrumentos"
	Case 2 TableName = "Area100Motores"
	Case Else TableName = ""
End Select

DateFrom = FormatDateToSQL($StartDateTime, 0)
$DateToSql = DateFrom

DateTo = FormatDateToSql(DateAdd("s", $Hour2Clock($StrGetElement($Duration,".",1)), DateFrom), 1)

Select Case $DataGrouping
'	Case 1 AvgGroupPeriod = "yy"
'	Case 2 AvgGroupPeriod = "q"	
'	Case 3 AvgGroupPeriod = "m"	
'	Case 4 AvgGroupPeriod = "dy"	
'	Case 5 AvgGroupPeriod = "dd"	
'	Case 6 AvgGroupPeriod = "ww"	
'	Case 7 AvgGroupPeriod = "dw"	
	Case 1 AvgGroupPeriod = "hh"	
	Case 2 AvgGroupPeriod = "mi"	
	Case Else AvgGroupPeriod = ""	
End Select

tagProgress = "Progress"

strSelectColumns = GetSelectedColumns(DBConnectionName, TableName)

Call ExportDBToCSV(TargetFile, DBConnectionName, TableName, DateFrom, DateTo, AvgGroupPeriod, tagProgress, strSelectColumns)

'TODO comentar
Function GetSelectedColumns(DBConnectionName, TableName)

	Dim sql, numCur, numRows, numCols, colName(), col, strColumns

	'select one row to get the column's names
	sql = "SELECT TOP 1 * FROM " & TableName 
	numCur = $DBCursorOpenSql(DBConnectionName, sql)
	numRows = $DBCursorRowCount(numCur)

	'if sql returns no result, exit function and return * (all columns)
	If numRows<1 Then
		$DBCursorClose(numCur)
		GetSelectedColumns = "*"
		Exit Function
	End If

	numCols = $DBCursorColumnCount(numCur)
	ReDim colName(numCols)

	For col=1 To numCols
		colName(col) = $DBCursorColumnInfo(numCur, $Num(col), 0)
	Next
	$DBCursorClose(numCur)


	'limit number of columns so it won't exceed an int (32bit) + 1st column (always present)
	limitCols = $Min(numCols, 33)

	For col=1 To numCols
		If col=1 Then
			strColumns = "[" & colName(col) & "]"

		'bit count goes from 0 to 31
		Else If $GetBit($intSQLColunas, col-2) = 1 Then
			strColumns = strColumns & ", [" & colName(col) & "]"

		End If
	Next
	
	GetSelectedColumns = strColumns
End Function