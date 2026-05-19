Function change_character_code(fileName, beforeCharacterCode, beforeLineSeparator, afterCharacterCode, afterLineSeparator)
    On Error Resume Next

    Dim beforeADODB :Set beforeADODB = CreateObject("ADODB.Stream")
    beforeADODB.Type = 2 ' adTypeText
    beforeADODB.Charset = beforeCharacterCode
    beforeADODB.Open
    beforeADODB.LineSeparator = ADODB_stream_lineseparator(beforeLineSeparator)
    beforeADODB.LoadFromFile fileName 'opne file

    Dim afterADODB :Set afterADODB = CreateObject("ADODB.Stream")
    afterADODB.Type = 2 ' adTypeText
    afterADODB.Charset = afterCharacterCode
    afterADODB.LineSeparator = ADODB_stream_lineseparator(afterLineSeparator)
    afterADODB.Open

    'If us use "CopyTo", the line separator code will remain as it is.
    Do Until beforeADODB.EOS
        line = beforeADODB.ReadText(-2)
        Call afterADODB.WriteText(line, 1)
    Loop

    beforeADODB.Close
    Set beforeADODB = Nothing

    Call afterADODB.SaveTofile(fileName, 2)
    afterADODB.Close
    Set afterADODB = Nothing

    If Err.Number <> 0 Then
        change_character_code = False
    Else
        change_character_code = True
    End If
End Function
