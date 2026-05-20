Function is_utf8(fileName)
    ' UTF-8（BOMあり・BOMなし問わず）かどうかを判定する
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  UTF-8ならTrue、それ以外はFalse

    If ADODB_stream_isCharcode(fileName, "utf-8") Then
        is_utf8 = True
    Else
        is_utf8 = False
    End If
End Function
