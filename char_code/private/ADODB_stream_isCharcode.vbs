Function ADODB_stream_isCharcode(fileName, charcode)
    On Error Resume Next
    Err.Clear

    Dim lowerCharcode
    Dim objBin, bytes, size
    Dim isUTF8BOM, isUTF16LEBOM, isUTF16BEBOM
    Dim isAscii, isUtf8Valid

    lowerCharcode = LCase(charcode)

    Set objBin = CreateObject("ADODB.Stream")
    objBin.Type = 1
    objBin.Open
    objBin.LoadFromFile fileName

    If Err.Number <> 0 Then
        ADODB_stream_isCharcode = False
        Err.Clear
        Exit Function
    End If

    size = objBin.Size
    If size > 0 Then
        bytes = objBin.Read(-1)
    Else
        bytes = ""
    End If

    objBin.Close
    Set objBin = Nothing

    If Err.Number <> 0 Then
        ADODB_stream_isCharcode = False
        Err.Clear
        Exit Function
    End If

    isUTF8BOM = False
    isUTF16LEBOM = False
    isUTF16BEBOM = False

    If size >= 3 Then
        If ADODB_stream_getByte(bytes, 0) = &HEF And ADODB_stream_getByte(bytes, 1) = &HBB And ADODB_stream_getByte(bytes, 2) = &HBF Then
            isUTF8BOM = True
        End If
    End If

    If (Not isUTF8BOM) And size >= 2 Then
        If ADODB_stream_getByte(bytes, 0) = &HFF And ADODB_stream_getByte(bytes, 1) = &HFE Then
            isUTF16LEBOM = True
        ElseIf ADODB_stream_getByte(bytes, 0) = &HFE And ADODB_stream_getByte(bytes, 1) = &HFF Then
            isUTF16BEBOM = True
        End If
    End If

    isAscii = True
    If size > 0 Then
        Dim i
        For i = 0 To size - 1
            If ADODB_stream_getByte(bytes, i) > 127 Then
                isAscii = False
                Exit For
            End If
        Next
    End If

    If isUTF8BOM Then
        isUtf8Valid = ADODB_stream_isValidUtf8(bytes, 3)
    Else
        isUtf8Valid = ADODB_stream_isValidUtf8(bytes, 0)
    End If

    Select Case lowerCharcode
        Case "ascii", "us-ascii"
            ADODB_stream_isCharcode = (Not isUTF8BOM) And (Not isUTF16LEBOM) And (Not isUTF16BEBOM) And isAscii
            Exit Function

        Case "utf-8-bom"
            ADODB_stream_isCharcode = isUTF8BOM And isUtf8Valid
            Exit Function

        Case "utf-8-nobom", "utf-8-no-bom"
            ADODB_stream_isCharcode = (Not isUTF8BOM) And (Not isUTF16LEBOM) And (Not isUTF16BEBOM) And isUtf8Valid
            Exit Function

        Case "utf-8", "utf8"
            ADODB_stream_isCharcode = (Not isUTF16LEBOM) And (Not isUTF16BEBOM) And isUtf8Valid
            Exit Function

        Case "unicode", "utf-16", "utf-16le", "utf-16-le"
            ADODB_stream_isCharcode = isUTF16LEBOM
            Exit Function

        Case "unicode big endian", "unicodefffe", "utf-16be", "utf-16-be"
            ADODB_stream_isCharcode = isUTF16BEBOM
            Exit Function

        Case "shift_jis", "shift-jis", "sjis", "cp932", "windows-31j"
            ADODB_stream_isCharcode = (Not isUTF8BOM) And (Not isUTF16LEBOM) And (Not isUTF16BEBOM) And (Not isUtf8Valid)
            Exit Function

        Case Else
            ADODB_stream_isCharcode = False
    End Select
End Function

Function ADODB_stream_isValidUtf8(bytes, startIndex)
    Dim i, n, b0, b1, b2, b3

    If IsEmpty(bytes) Then
        ADODB_stream_isValidUtf8 = True
        Exit Function
    End If

    n = LenB(bytes) - 1
    If n < 0 Then
        ADODB_stream_isValidUtf8 = True
        Exit Function
    End If

    i = startIndex

    Do While i <= n
        b0 = ADODB_stream_getByte(bytes, i)

        If b0 <= &H7F Then
            i = i + 1

        ElseIf b0 >= &HC2 And b0 <= &HDF Then
            If i + 1 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            If b1 < &H80 Or b1 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 2

        ElseIf b0 = &HE0 Then
            If i + 2 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            If b1 < &HA0 Or b1 > &HBF Or b2 < &H80 Or b2 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 3

        ElseIf b0 >= &HE1 And b0 <= &HEC Then
            If i + 2 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            If b1 < &H80 Or b1 > &HBF Or b2 < &H80 Or b2 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 3

        ElseIf b0 = &HED Then
            If i + 2 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            If b1 < &H80 Or b1 > &H9F Or b2 < &H80 Or b2 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 3

        ElseIf b0 >= &HEE And b0 <= &HEF Then
            If i + 2 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            If b1 < &H80 Or b1 > &HBF Or b2 < &H80 Or b2 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 3

        ElseIf b0 = &HF0 Then
            If i + 3 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            b3 = ADODB_stream_getByte(bytes, i + 3)
            If b1 < &H90 Or b1 > &HBF Or b2 < &H80 Or b2 > &HBF Or b3 < &H80 Or b3 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 4

        ElseIf b0 >= &HF1 And b0 <= &HF3 Then
            If i + 3 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            b3 = ADODB_stream_getByte(bytes, i + 3)
            If b1 < &H80 Or b1 > &HBF Or b2 < &H80 Or b2 > &HBF Or b3 < &H80 Or b3 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 4

        ElseIf b0 = &HF4 Then
            If i + 3 > n Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            b1 = ADODB_stream_getByte(bytes, i + 1)
            b2 = ADODB_stream_getByte(bytes, i + 2)
            b3 = ADODB_stream_getByte(bytes, i + 3)
            If b1 < &H80 Or b1 > &H8F Or b2 < &H80 Or b2 > &HBF Or b3 < &H80 Or b3 > &HBF Then
                ADODB_stream_isValidUtf8 = False
                Exit Function
            End If
            i = i + 4

        Else
            ADODB_stream_isValidUtf8 = False
            Exit Function
        End If
    Loop

    ADODB_stream_isValidUtf8 = True
End Function

Function ADODB_stream_getByte(bytes, idx)
    ADODB_stream_getByte = AscB(MidB(bytes, idx + 1, 1)) And &HFF
End Function
