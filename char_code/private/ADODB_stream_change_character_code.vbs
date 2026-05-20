Function ADODB_stream_change_character_code(fileName, fromCharcode, toCharcode, addBOM)
    On Error Resume Next
    Err.Clear

    Dim adoFrom
    Select Case LCase(fromCharcode)
        Case "utf-8-bom", "utf-8-nobom", "utf-8-no-bom", "utf8"
            adoFrom = "utf-8"
        Case "ascii", "us-ascii"
            adoFrom = "us-ascii"
        Case Else
            adoFrom = fromCharcode
    End Select

    Dim srcStream
    Set srcStream = CreateObject("ADODB.Stream")
    srcStream.Type = 2
    srcStream.Charset = adoFrom
    srcStream.Open
    srcStream.LoadFromFile fileName

    If Err.Number <> 0 Then
        ADODB_stream_change_character_code = False
        Err.Clear
        Exit Function
    End If

    Dim text
    text = srcStream.ReadText(-1)

    srcStream.Close
    Set srcStream = Nothing

    If Err.Number <> 0 Then
        ADODB_stream_change_character_code = False
        Err.Clear
        Exit Function
    End If

    Dim dstStream
    Set dstStream = CreateObject("ADODB.Stream")
    dstStream.Type = 2
    dstStream.Charset = toCharcode
    dstStream.Open

    If addBOM Then
        dstStream.WriteText ChrW(65279) & text, 0
    Else
        dstStream.WriteText text, 0
    End If

    dstStream.SaveToFile fileName, 2
    dstStream.Close
    Set dstStream = Nothing

    If Err.Number = 0 Then
        ADODB_stream_change_character_code = True
    Else
        ADODB_stream_change_character_code = False
        Err.Clear
    End If
End Function
