' * Description: Returns the format used for date
' * Arguments:
'  - DateSource: 0=Studio; 1=VBScript
' * Returned Value: Format used for date (e.g.: MDY, DMY, etc)
' * Dependences: GetDateSeparator()
Function GetDateFormat(DateSource)
	Dim factorRef, dtRef, separator, dateOnly, dt(3), id, fmt(3), dtSource
	If $Day>1 Then
		factorRef = -1
	Else
		factorRef = 1
	End If
	If DateSource=1 Then
		dtSource = Date
		dtRef = DateAdd("d", factorRef, Date)
	Else
		dtSource = $Date
		dtRef = $ClockGetDate(60*60*24*factorRef+ $DateTime2Clock( $Date , "00:00:00"))
	End If
	separator = GetDateSeparator(DateSource)
	dateOnly = $StrGetElement(dtSource, " ", 1)
	dt(1) = $StrGetElement(dateOnly, separator, 1)
	dt(2) = $StrGetElement(dateOnly, separator, 2)
	dt(3) = $StrGetElement(dateOnly, separator, 3)
	For 	id=1 To 3
		dt(0) = $StrGetElement(dtRef, separator, id)
		If dt(id)<>dt(0) Then
			fmt(id) = "D"
		ElseIf Len(dt(id))=4 Then
			fmt(id) = "Y"
		Else
			fmt(id) = "M"	
		End If	
	Next
	GetDateFormat = fmt(1) & fmt(2) & fmt(3)
End Function


' * Description: Returns the separator used for date in Studio
' * Arguments:
'  - DateSource: 0=Studio; 1=VBScript
' * Returned Value: Separator used for date in Studio (e.g.: "/")
' * Dependences: None
Function GetDateSeparator(DateSource)
	Dim charPos, charStr, dtSource
	If DateSource=1 Then
		dtSource = Date
	Else
		dtSource = $Date
	End If
	For charPos=1 To Len(dtSource)
		charStr = Mid(dtSource, charPos, 1)
		If IsNumeric(charStr)=False Then Exit For
	Next
	GetDateSeparator = charStr
End Function


' * Description: Returns the timestamp in the military time format (instead of AM/PM)
' * Arguments: 
'  - TimeStampAMPM: TimeStamp With the time in the AM/PM format
' * Returned Value: Timestamp in the the military time format (instead of AM/PM)
' * Dependences: None
Function FormatTimeToMil(TimeAMPM)
	Dim ampm, hh, mm, ss, hhMil
	hh = $StrGetElement(TimeAMPM, ":", 1)
	mm = $StrGetElement(TimeAMPM, ":", 2)
	ss = $StrGetElement($StrGetElement(TimeAMPM, ":", 3), " ", 1)
	ampm = $StrGetElement(TimeAMPM, " ", 2)
	If ampm="" Then 'TimeAMPM already in military format
		hhMil = hh
	ElseIf UCase(ampm)="AM" Then 'AM
		If $Num(hh)=12 Then
			hhMil = 0
		Else
			hhMil = hh
		End If
	Else  'PM
		If $Num(hh)=12 Then
			hhMil = hh
		Else
			hhMil = hh+12
		End If
	End If
	FormatTimeToMil = $Format("%02s", hhMil) & ":" & $Format("%02s", mm) & ":" & $Format("%02s", ss)
End Function


' * Description: Returns the timestamp in the format supported by SQL Server
' * Arguments: 
'  - TimeStamp: TimeStamp With the Date In the current format Set For Studio
'  - DateSource: 0=Studio; 1=VBScript
' * Returned Value: Timestamp in the format supported by SQL Server (YYYY-MM-DD HH:MM:SS)
' * Dependences: GetDateSeparator(), GetDateFormat()
Function FormatDateToSQL(TimeStamp, DateSource)
	Dim dateOnly, timeOnly, dtSeparator, dtFormat, dt(3), id, refFormat, ampm
	dateOnly = $StrGetElement(TimeStamp, " ", 1)
	timeOnly = $StrGetElement(TimeStamp, " ", 2)
	ampm = $StrGetElement(TimeStamp, " ", 3)
	If ampm<>"" Then 'ampm time format
		timeOnly = FormatTimeToMil(timeOnly & " " & ampm)
	Else 'military time format
		timeOnly = $StrGetElement(timeOnly, ".", 1)
	End If	
	dtSeparator = 	GetDateSeparator(DateSource)
	dtFormat = GetDateFormat(DateSource)
	For id=1 To 3
		refFormat = UCase(Mid(dtFormat, id, 1))
		dt(0) = $StrGetElement(dateOnly, dtSeparator, id)
		If refFormat="Y" Then
			dt(1) = dt(0)
		ElseIf refFormat="M" Then
			dt(2) = dt(0)
		Else
			dt(3) = dt(0)
		End If
	Next
	FormatDateToSQL = RTrim(dt(1) & "-" & dt(2) & "-" & dt(3) & " " & timeOnly)
End Function



