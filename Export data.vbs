'Variables available only for this group can be declared here.

Dim TargetFile, DBConnectionName, TableName, DateFrom, DateTo, AvgGroupPeriod, tagProgress

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

Call ExportDBToCSV(TargetFile, DBConnectionName, TableName, DateFrom, DateTo, AvgGroupPeriod, tagProgress, $intSQLColunas)