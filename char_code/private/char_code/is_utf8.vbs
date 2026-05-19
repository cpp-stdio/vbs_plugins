Function is_utf8(fileName)
    ' The character encoding for this file is UTF-8?
    '
    'Parameters
    '----------
    'fileName : String
    '   It is the file name whose character code us want to change
    '
    'Return
    '----------
    'boolen
    '   Yes(True) , No(False)

    WScript.Echo fileName

    If ADODB_stream_isCharcode(fileName, "UTF-8") Then
        WScript.Echo "[SUCCESS] This text file character is UTF-8."
        is_utf8 = True
    Else
        WScript.Echo "[FAILURE] This text file character isn't UTF-8."
        is_utf8 = False
    End If
End Function
