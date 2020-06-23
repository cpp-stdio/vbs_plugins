Function ADODB_Stream_lineSeparator(lineSeparator)

	If StrComp(lineSeparator, vbCr, 0) = 0 Then
		ADODB_Stream_lineSeparator = 13 'carriage return
		Exit Function
	End If
	
	If StrComp(lineSeparator, vbLf, 0) = 0 Then
		ADODB_Stream_lineSeparator = 10 'line feed
		Exit Function
	End If

	ADODB_Stream_lineSeparator = -1 'carriage return and line feed(ADODB.Stream default)

End Function