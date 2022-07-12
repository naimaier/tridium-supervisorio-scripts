' * Description: Returns the Catalog Name from the given Connection String
' * Arguments: 
'  - strConnectionString: Connection String value
' * Returned Value: Catalog Name extracted from the given Connection String
' * Dependences: None
Function GetCatalogName(strConnectionString)
	Dim arr,id
	arr = Split(strConnectionString, ";")
	For id=0 To UBound(arr)
		If InStr(LCase(Replace(arr(id), " ", "")), "initialcatalog=")>0 Then
			GetCatalogName = $StrGetElement(arr(id), "=", 2)
			Exit For
		End If
	Next
End Function

' * Description: Remove the Catalog Name instruction from a given Connection String
' * Arguments: 
'  - strConnectionString: Connection String value
' * Returned Value: Connection String without the Catalog Name instruction
' * Dependences: None
Function RemoveCatalogName(strConnectionString)
	Dim arr,id, res
	arr = Split(strConnectionString, ";")
	res = strConnectionString
	For id=0 To UBound(arr)
		If InStr(LCase(Replace(arr(id), " ", "")), "initialcatalog=")>0 Then
			res = Replace(Replace(strConnectionString, arr(id), ""), ";;", ";")
			Exit For
		End If
	Next
	RemoveCatalogName = res
End Function

' * Description: Creates a new database (catalog)
' * Arguments: 
'  - strConnectionName: Name of the Database/ERP connection created in the "Tasks > Database/ERP interface. The default value is DB"
'  - strDatabaseName: Name of the database (catalog) that must be created"
' * Returned Value: This function returns the total number of rows affected by the SQL statement. If an error occurs, then it returns a negative number.
' * Dependences: None
Function CreateDatabase(strConnectionName, strDatabaseName)
	Dim sql
	If strConnectionName = "" Then strConnectionName = "DB"
	sql = "CREATE DATABASE " & strDatabaseName
	If strDatabaseName<>"" Then CreateDatabase = $DBExecute(strConnectionName, sql)
End Function

' * Description: Checks if a datatabase (catalog) exists. If not, attempts to create it.
' * Arguments: 
'  - strConnectionName: Name of the Database/ERP connection created in the "Tasks > Database/ERP interface. The default value is DB"
'  - strConnectionStringTag: Name of the tag configured in the connection string field of the Database/ERP connection created in the "Tasks > Database/ERP interface"
' * Returned Value: 0 = Error (database not available); 1 = Success (database available)
' * Dependences: GetCatalogName, RemoveCatalogName, CreateDatabase
Function InitDB(strConnectionName, strConnectionStringTag)
	Dim strConnectionString, CatalogName, sql, numCur, numRows, ret
	If strConnectionName="" Or $GetTagValue(strConnectionStringTag)="" Then 
		ret = 0
	Else
		strConnectionString = $GetTagValue(strConnectionStringTag)
		CatalogName = GetCatalogName(strConnectionString)
		$SetTagValue(strConnectionStringTag, RemoveCatalogName(strConnectionString))
		sql = "SELECT * FROM sys.databases WHERE name='" & CatalogName & "'"
		numCur = $DBCursorOpenSQL(strConnectionName, sql)
		numRows = $DBCursorRowCount(numCur)
		$DBCursorClose(numCur)
		If numRows>=1 Then
			ret = 1
		Else	
			Call CreateDatabase(strConnectionName, CatalogName)
			numCur = $DBCursorOpenSQL(strConnectionName, sql)
			numRows = $DBCursorRowCount(numCur)
			$DBCursorClose(numCur)
			If numRows>=1 Then 
				ret = 1
			Else
				ret = 0
			End If		
		End If
		$SetTagValue(strConnectionStringTag, strConnectionString)
	End If	
	InitDB = ret
End Function


