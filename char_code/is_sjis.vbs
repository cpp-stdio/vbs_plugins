Function is_sjis(fileName)
    ' The character encoding for this file is SJIS(SHIFT-JIS)?
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

    If ADODB_stream_isCharcode(fileName, "Shift_JIS") Then
        WScript.Echo "[SUCCESS] This text file character is Shift_JIS."
        is_sjis = True
    Else
        WScript.Echo "[FAILURE] This text file character isn't Shift_JIS."
        is_sjis = False
    End If
End Function
