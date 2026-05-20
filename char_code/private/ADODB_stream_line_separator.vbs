Function ADODB_stream_line_separator(lineSeparator)
    If StrComp(lineSeparator, vbCr, 0) = 0 Then
        ADODB_stream_line_separator = 13
        Exit Function
    End If
    
    If StrComp(lineSeparator, vbLf, 0) = 0 Then
        ADODB_stream_line_separator = 10
        Exit Function
    End If

    ADODB_stream_line_separator = -1

End Function
