Function utf8_for_sjis(fileName, beforeLineSeparator, afterLineSeparator)
    'This program change character code
    ' UTF-8 Å® SJIS(SHIFT-JIS) 
    '
    'Parameters
    '----------
    'fileName : String
    '   It is the file name whose character code us want to change
    'beforeLineSeparator : String
    '   It is before change of line separator code.
    '   The line separator are as follows three patterns.
    '   [vbCrLf] : carriage return and line feed
    '   [ vbCr ] : carriage return
    '   [ vbLf ] : line feed
    'afterLineSeparator : String
    '   It is after change of line separator code.
    '   The line separator are as follows three patterns.
    '   [vbCrLf] : carriage return and line feed
    '   [ vbCr ] : carriage return
    '   [ vbLf ] : line feed
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    
    If change_character_code(fileName, "UTF-8", beforeLineSeparator, "Shift_JIS", afterLineSeparator) Then
        WScript.Echo "[SUCCESS] change character code. UTF-8 Å®Å@SJIS"
        sjis_for_utf8 = True
    Else
        WScript.Echo "[FAILURE] change character code. UTF-8Å@Å®Å@SJIS"
        sjis_for_utf8 = False
    End if
End Function