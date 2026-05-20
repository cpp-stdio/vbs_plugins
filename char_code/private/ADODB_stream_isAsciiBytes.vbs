Function ADODB_stream_isAsciiBytes(fileName)
    On Error Resume Next
    Err.Clear

    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objStream.LoadFromFile fileName

    If Err.Number <> 0 Then
        ADODB_stream_isAsciiBytes = False
        Err.Clear
        Exit Function
    End If

    If objStream.Size = 0 Then
        objStream.Close
        Set objStream = Nothing
        ADODB_stream_isAsciiBytes = True
        Exit Function
    End If

    Dim allBytes
    allBytes = objStream.Read(-1)

    objStream.Close
    Set objStream = Nothing

    If Err.Number <> 0 Then
        ADODB_stream_isAsciiBytes = False
        Err.Clear
        Exit Function
    End If

    Dim i
    For i = 0 To UBound(allBytes)
        If allBytes(i) > 127 Then
            ADODB_stream_isAsciiBytes = False
            Exit Function
        End If
    Next

    ADODB_stream_isAsciiBytes = True
End Function
