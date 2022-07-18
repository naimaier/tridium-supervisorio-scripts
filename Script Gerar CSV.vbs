Option Explicit
'Keep the Option Explicit statement in the first line of this interface.

'Procedures with global scope can be implemented here.
'Global variables are NOT supported in this interface.


' * Description: Exports data from a SQL Server database table into a CSV file
' * Arguments:
'  - TargetFile: Path and name of the target CSV file (e.g.: "c:\temp\Test.csv").
' - DBConnectionName: Name (alias) of the 'Database/ERP' connection (e.g.: "DB")
' - TableName: Name of the table where the values will be imported from.
' - DateFrom: Initial date for the filtered data from the database. If omitted, the entire table will be exported (up to 16000 records)
' - DateTo: Final date for the filtered data from the database. If omitted, the entire table will be exported (up to 16000 records)
' - AvgGroupPeriod: Type of grouping for the records. If omitted, the actual original records will be exported. Otherwise, the average for the values will be calculated based on the grouping period:
'   year: "yy", quarter: "q", month: "m", dayofyear: "dy", day: "dd", week: "ww", weekday: "dw", hour: "hh", minute: "mi")
' - tagProgress: String value with the name of the tag that will receive the progress value (0-100%) for the values exported to Excel. If omitted (""), the progress will not be tracked.
' - UserSelectedColumns: An integer whose bits represents the columns selected by the user
'     with it's least significant bit representing the second column of the table 
'     (e.g.: 0101 -> selects columns 2 and 4) + the first column is always present
' * Returned Value: Number of rows exported to the CSV file. 
' * Dependences: None
Function ExportDBToCSV(TargetFile, DBConnectionName, TableName, DateFrom, DateTo, AvgGroupPeriod, tagProgress, UserSelectedColumns)

	Dim sql, numRows, row, numCur, txt, numCols, col, colName(), prefix, suffix, prefixTime_Stamp, sqlCols, strQueryColumns

	'make sure the file extension is "csv"
	If InStr(TargetFile, ".csv")<1 Then TargetFile = TargetFile & ".csv" 

	strQueryColumns = GetSelectedColumnsQuery(DBConnectionName, TableName, UserSelectedColumns)
	
	'select one row to get the column's names
	sql = "SELECT TOP 1 " & strQueryColumns & " FROM " & TableName 
	numCur = $DBCursorOpenSql(DBConnectionName, sql)
	numRows = $DBCursorRowCount(numCur)
	If numRows>0 Then
			numCols = $DBCursorColumnCount(numCur)
			ReDim colName(numCols)
			txt = ""
			sqlCols = ""
			For col=1 To numCols
				colName(col) = $DBCursorColumnInfo(numCur, $Num(col), 0)
				txt = txt & colName(col)
				'create sql statement necessary when grouping values (necessary only if AvgGroupPeriod <> "")
				If col>1 Then
					sqlCols = sqlCols & "Avg(" & colName(col) & ") AS " & colName(col)
				Else
					sqlCols = sqlCols & "Min(" & colName(col) & ") AS " & colName(col)
				End If
				If col<numCols Then 
					txt = txt & ";"
					sqlCols = sqlCols & ","
				End If
			Next
			'creates the header for the target file
			$FileWrite(TargetFile, txt, 0)
	End If
	$DBCursorClose(numCur)
	
	'queries the actual data to be exported to the csv file
	sql = ""
	If (DateFrom<>"" And DateTo<>"") Then sql = " WHERE Time_Stamp>='" & DateFrom & "' AND Time_Stamp<='" & DateTo & "'"
	If AvgGroupPeriod="" Then 
		sql = "SELECT " & strQueryColumns & " FROM " & TableName & sql & " ORDER BY Time_Stamp"
	Else
		sql = "SELECT " & sqlCols & " FROM " & TableName & sql & " GROUP BY DateAdd(" & AvgGroupPeriod & ",DateDiff(" & AvgGroupPeriod & ",0,Time_Stamp),0)" & " ORDER BY Min(Time_Stamp)"
	End If	
	numCur = $DBCursorOpenSQL(DBConnectionName, sql)
	numRows = $DBCursorRowCount(numCur)

	For row=1 To numRows
		txt = ""
		For col=1 To numCols
			If col>1 Then txt = txt & ";"
			txt = txt & $DBCursorGetValue(numCur, colName(col))
		Next
		$FileWrite(TargetFile, txt, 1)
		$DBCursorNext(numCur)
		'updates the tag (if any) used to track the progress of the operation
		If tagProgress<>"" Then $SetTagValue(tagProgress, row/numRows*100)
	Next
	$DBCursorClose(numCur)
	
	'returns the number of rows actually exported to the target file
	ExportDBToCSV = $max(0, numRows)
End Function


'DESCRIPTION: Generates the SQL query's column selection part containing the columns selected by the user.
'PARAMETERS:
' - DBConnectionName: Name (alias) of the 'Database/ERP' connection (e.g.: "DB")
' - TableName: Name of the table where the values will be imported from.
' - UserSelectedColumns: An integer whose bits represents the columns selected by the user
'     with it's least significant bit representing the second column of the table 
'     (e.g.: 0101 -> selects columns 2 and 4) + the first column is always present
'RETURNED VALUES: A string containing a SQL query's column selection part
'DEPENDENCIES: None
'AUTHOR: Henrique Morin Naimaier
Function GetSelectedColumnsQuery(DBConnectionName, TableName, UserSelectedColumns)

	Dim sql, numCur, numRows, numCols, limitCols, colName(), col, strQueryColumns

	'select one row to get the column's names
	sql = "SELECT TOP 1 * FROM " & TableName 
	numCur = $DBCursorOpenSql(DBConnectionName, sql)
	numRows = $DBCursorRowCount(numCur)

	'if sql returns no result, exit function and return * (all columns)
	If numRows<1 Then
		$DBCursorClose(numCur)
		GetSelectedColumnsQuery = "*"
		Exit Function
	End If

	'get the column's names
	numCols = $DBCursorColumnCount(numCur)
	ReDim colName(numCols)

	For col=1 To numCols
		colName(col) = $DBCursorColumnInfo(numCur, $Num(col), 0)
	Next
	$DBCursorClose(numCur)


	'limit number of columns so it won't exceed 33,
	'33: an int (32bit) 32 selectable columns + 1st column (timestamp, always present)
	limitCols = $Min(numCols, 33)

	'fill the string with selected columns' names
	For col=1 To limitCols
		If col=1 Then
			strQueryColumns = "[" & colName(col) & "]"

		'bit count goes from 0 to 31
		ElseIf $GetBit(UserSelectedColumns, col-2) = 1 Then
			strQueryColumns = strQueryColumns & ", [" & colName(col) & "]"

		End If
	Next
	
	GetSelectedColumnsQuery = strQueryColumns
End Function


//$region: Report Functions

'DESCRIPTION: Writes data to the file. This function must be executed on the Server to create the file on the Server computer
'PARAMETERS:
' fileName: Name of the report file
' txt: Data to be saved into the report file
' append: 0=Overwrites the file; 1=Append the file
'RETURNED VALUES: None
'DEPENDENCIES: None
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepWriteToFile(fileName, txt, append)
	txt = Replace(txt, $Asc2Str(13), vbCrLf)
	txt = Replace(txt, $Asc2Str(27), """")
	$FileWrite(fileName, txt, $Num(append))
End Sub



'DESCRIPTION: Creates the stylesheet (CSS file) for the report in the \Web sub-folder of the application. This function must be executed on the Server.
'PARAMETERS: None
'RETURNED VALUES: None
'DEPENDENCIES: None
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepCreateCSS()
	Dim txt, fileName
	fileName = $GetAppPath() & "Web\Rep.css"
	If $FindFile(fileName)<>1 Then
		txt = ".headertable{border: none; border-bottom:double gray}"
		txt = txt & vbCrLf & ".header{border:none}"
		txt = txt & vbCrLf &  ".title1{font-size: large}"
		txt = txt & vbCrLf &  ".title2{font-size: medium}"
		txt = txt & vbCrLf &  ".title3{font-size: small}"
		txt = txt & vbCrLf &  ".bg {background: LightGrey}"
		txt = txt & vbCrLf &  ".portrait {padding:36pt; background:White; width:8.5in; height:11in; margin: 1% auto; page-break-after:always}"
		txt = txt & vbCrLf &  ".landscape {padding:36pt; background:White; width:11in; height:8.5in; margin: 1% auto; page-break-after:always}"
		txt = txt & vbCrLf &  "h2 {font-family:Verdana;  font-weight:bold}"
		txt = txt & vbCrLf &  "h3 {font-family:Verdana;  font-weight:bold}"
		txt = txt & vbCrLf &  "table {border:double gray; width=100%; text-align=center}"
		txt = txt & vbCrLf &  "th {border:1px solid;font-family:Verdana; font-size=x-small; background-color:beige}"
		txt = txt & vbCrLf &  "td {border:1px solid;font-family:Verdana; font-size=x-small}"
		txt = txt & vbCrLf &  "p {font-family:Verdana; font-size=x-small; margin:0.5em}"
		$FileWrite(fileName, txt, 0)
	End If
End Sub

'DESCRIPTION: Writes the Header of the report
'PARAMETERS: None
'RETURNED VALUES: None
'DEPENDENCIES: None
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepBuildHeader()
	Dim FileName, Orientation, Title1, Title2, Title3, Logo, txt
	FileName = $Rep.FileName
	Orientation = $Rep.Orientation
	Title1 = $Rep.Title1
	Title2 = $Rep.Title2
	Title3 = $Rep.Title3
	Logo = $Rep.Logo
	txt = "<table class=headertable>"
	txt = txt & vbCrLf & "<tr>"
	If Logo<>"" Then txt = txt & vbCrLf & "<td Class=""header title1"" rowspan=3 width=144pt><img src=""" & Logo & """ height=72pt></td>"
	txt = txt & vbCrLf & "<td class=""header title1"">" & Title1 & "</td>"
	txt = txt & vbCrLf & "</tr>"
	txt = txt & vbCrLf & "<tr>"
	txt = txt & vbCrLf & "<td class=""header title2"">" & Title2 & "</td>"
	txt = txt & vbCrLf & "</tr>"
	txt = txt & vbCrLf & "<tr>"
	txt = txt & vbCrLf & "<td class=""header title3"">" & Title3 & "</td>"
	txt = txt & vbCrLf & "</tr>"
	txt = txt & vbCrLf & "</table>"
	txt = Replace(txt, vbCrLf, $Asc2Str(13))
	txt = Replace(txt, """", $Asc2Str(27))
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
End Sub

'DESCRIPTION: Breaks the current page and create a header for the next page, if applicable
'PARAMETERS: None
'RETURNED VALUES: None
'DEPENDENCIES:
' RepBuildHeader()
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepBreakPage()
	Dim FileName, RepeatHeader, Orientation, txt
	FileName = $Rep.FileName
	RepeatHeader = $Rep.RepeatHeader
	Orientation = $Rep.Orientation
	txt = "</div>"
	txt = txt & vbCrLf & "<div class=" & Orientation & ">"
	txt = Replace(txt, vbCrLf, $Asc2Str(13))
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
	If RepeatHeader=1 Then
		Call RepBuildHeader()
	End If
	$Rep.NumRow = 0
	$Rep.NumPage = $Rep.NumPage+1
End Sub

'DESCRIPTION: Checks if the current page is filled, so a page break must be created
'PARAMETERS: None
' rowIncrement: Number of rows that should be incremented in the current page by the last information appended into the report
' dataEnd: Data that must be appended at the end of the current page before breaking it
' dataStart: Data that must be appended to the beggining of the next page after breaking the previous page
'RETURNED VALUES: None
'DEPENDENCIES:
' RepBreakPage()
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepCheckPageBreak(rowIncrement, dataEnd, dataStart)
	Dim maxRows, maxPortrait, maxLandscape, Orientation, FileName, RepeatHeader, numPage
	Orientation = $Rep.Orientation
	FileName = $Rep.FileName
	RepeatHeader = $Rep.RepeatHeader
	numPage = $Rep.NumPage
	maxPortrait = 38
	maxLandscape = 22
	If LCase(Orientation)="landscape" Then
		maxRows = maxLandscape
	Else
		maxRows = maxPortrait
	End If
	If (RepeatHeader=0 And numPage>1) Then maxRows = maxRows +3
	$Rep.NumRow = $Rep.NumRow+rowIncrement
	If $Rep.NumRow>maxRows Then 
		If dataEnd<>"" Then $RunGlobalProcedureOnServer("RepWriteToFile", FileName, dataEnd, 1)
		Call RepBreakPage()
		If dataStart<>"" Then $RunGlobalProcedureOnServer("RepWriteToFile", FileName, dataStart, 1)
	End If
End Sub

'DESCRIPTION: Starts a new report. This function must be called once when generating a new report
'PARAMETERS: None
' FileName: Name of the output report file (without the path). The report will be automatically created in the \Web sub-folder of the application
' Orientation: Either "Portrait" or "Landscape"
' RepeatHeader: 0=The Header is created only in the first page; 1=The Header is created in each page
' Title1: Text in the first line of the Header
 'Title2: Text in the second line of the Header
 'Title3: Text in the third line of the Header
 'Logo: Name of the picture displayed in the Header
'RETURNED VALUES: None
'DEPENDENCIES: None
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepStart(FileName, Orientation, RepeatHeader, Title1, Title2, Title3, Logo)
	Dim txt
	'Move settings to global tags
	If InStr(FileName, ".html")=0 Then FileName = FileName & ".html"
	FileName = $GetAppPath() & "Web\" & FileName
	$Rep.FileName = FileName
	If LCase(Orientation)<>"landscape" Then Orientation = "portrait"
	$Rep.Orientation = Orientation
	$Rep.RepeatHeader = RepeatHeader
	$Rep.Title1 = Title1
	$Rep.Title2 = Title2
	$Rep.Title3 = Title3
	$Rep.Logo = Logo
	$Rep.NumRow = 0
	$Rep.NumPage = 1
	'Create the HTML file
	$RunGlobalProcedureOnServer("RepCreateCSS")
	txt = "<head>"
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 0)
	txt = "<link rel=stylesheet type=text/css href=Rep.css />"
	txt = txt & vbCrLf & "</head>"
	txt = txt & vbCrLf & "<body class=bg>"
	txt = txt & vbCrLf & "<div class=" & Orientation & ">"
	txt = Replace(txt, vbCrLf, $Asc2Str(13))
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
	Call RepBuildHeader()
End Sub

'DESCRIPTION: Appends a text into the report
'PARAMETERS:
' text: Text to the appended to the report
'RETURNED VALUES: None
'DEPENDENCIES:
' RepCheckPageBreak()
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepAppendText(text)
	Dim FileName, txt
	FileName = $Rep.FileName
	txt = "<p>" & text & "</p>"
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
	Call RepCheckPageBreak(2, "", "")
End Sub

'DESCRIPTION: Appends a image into the report
'PARAMETERS:
' text: Text to the appended to the report
' image: Image to the appended to the report
'RETURNED VALUES: None
'DEPENDENCIES:
' RepCheckPageBreak()
'AUTHOR: Marcelo Naimaier
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepAppendImage(Image, ImageLegenda)
	Dim FileName, txt
	FileName = $Rep.FileName
	txt = txt & vbCrLf & "<tr>"
	If Image<>"" Then txt = txt & vbCrLf & "<img src=""" & Image & """ width=1050pt height=630pt></td>"
	txt = txt & vbCrLf & "</tr>"
	txt = txt & vbCrLf & "<tr>"
	If ImageLegenda<>"" Then txt = txt & vbCrLf & "<img src=""" & ImageLegenda & """ width=1050pt height=373pt></td>"
	txt = txt & vbCrLf & "</tr>"
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
	Call RepCheckPageBreak(2, "", "")
End Sub


'DESCRIPTION: Appends a table (from a database) into the report
'PARAMETERS:
' DBConnection: Name of the database connection used to query data from the database (from Tasks > External Database > Connections)
' CSVLabels: Labels used in the header of the table, separated by comma (e.g.: "Label1,Label2,Label3"). If omitted, all fields from the table will be retrieved
' CSVFields: Fields from the database. If omitted, the labels will be used as the field names. If the CSVLabels field is blank, this parameter is ignored.
' Table: Name of the database table where the values must be retrieved from
' Condition: Filter condition to query data from the database. If omitted, all records from the table will be retrieved
'RETURNED VALUES: None
'DEPENDENCIES:
' RepCheckPageBreak()
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepAppendTable(DBConnection, CSVLabels, CSVFields, Table, Condition)
	Dim sql, field, numFields, fieldName(), labelName(), numCur, numRows, row, FileName, tableHeader, txt
	FileName = $Rep.FileName
	If CSVLabels="" Then 'select the whole table (the user did NOT specify the fields manually)
		sql = "SELECT * FROM " & Table
		If Condition<>"" Then sql = sql & " WHERE " & Condition
		numCur = $DBCursorOpenSQL(DBConnection, sql)
		numFields = $DBCursorColumnCount(numCur)
		If numFields>0 Then
			ReDim Preserve labelName(numFields)
			ReDim Preserve fieldName(numFields)
		End If
		For field=1 To numFields
			labelName(field) = $DBCursorColumnInfo(numCur, $Num(field), 0)
			fieldName(field) = labelName(field)
		Next
	Else 'the user specified the fields manually
		numFields = (Len(CSVLabels) - Len(Replace(CSVLabels, ",", ""))) + 1
		ReDim Preserve labelName(numFields)
		ReDim Preserve fieldName(numFields)
		sql = "SELECT "
		For field=1 To numFields
			labelName(field) = $StrGetElement(CSVLabels, ",", field)
			fieldName(field) = $StrGetElement(CSVFields, ",", field)
			If fieldName(field)="" Then fieldName(field) = labelName(field)
			If field<>1 Then sql = sql & ","
			sql = sql & fieldName(field)
		Next
		sql = sql & " FROM " & Table
		If Condition<>"" Then sql = sql & " WHERE " & Condition
		numCur = $DBCursorOpenSQL(DBConnection, sql)
	End If
	numRows = $DBCursorRowCount(numCur)
	For row=1 To numRows
		If row=1 Then 'table header
			txt = "<table>"
			txt = txt & vbCrLf & "<tr>"
			For field=1 To numFields
				txt = txt & vbCrLf & "<th>" & labelName(field) & "</th>"
			Next
			txt = txt & vbCrLf & "</tr>"
			txt = Replace(txt, vbCrLf, $Asc2Str(13))
			$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
			tableHeader = txt
			Call RepCheckPageBreak(1, "</table>", txt)
		End If
		txt = "<tr>"
		For field=1 To numFields
			txt = txt & vbCrLf & "<td>" & $DBCursorGetValue(numCur, fieldName(field)) & "</td>"
		Next
		txt = txt & vbCrLf & "</tr>"
		txt = Replace(txt, vbCrLf, $Asc2Str(13))
		$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
		Call RepCheckPageBreak(1, "</table>", tableHeader)		
		If row=numRows Then
			txt = "</table>"
			$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
		End If
		$DBCursorNext(numCur)
	Next
	$DBCursorClose(numCur)
End Sub

'DESCRIPTION: Finishes the report file. This function must be called once after appending all data into the report.
'PARAMETERS: None
'RETURNED VALUES: None
'DEPENDENCIES: None
'AUTHOR: Fabio Terezinho
'Revisions: 
' 1.0.0.0: Initial revision
Sub RepEnd()
	Dim FileName, txt
	FileName = $Rep.FileName
	txt = "</div>"
	txt = txt & vbCrLf & "</body>"
	txt = Replace(txt, vbCrLf, $Asc2Str(13))
	$RunGlobalProcedureOnServer("RepWriteToFile", FileName, txt, 1)
End Sub


