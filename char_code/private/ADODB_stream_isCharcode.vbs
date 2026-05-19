Function ADODB_stream_isCharcode(fileName, charcode)
    On Error Resume Next

    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 'Text
    objStream.Charset = charcode
    objStream.Open
    objStream.LoadFromFile fileName

    Dim text
    text = objStream.ReadText

    objStream.Close
    Set objStream = Nothing

    If Err.Number = 0 Then
        ADODB_stream_isCharcode = True
    Else
        ADODB_stream_isCharcode = False
        Err.Clear
    End If
End Function
