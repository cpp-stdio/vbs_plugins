Function is_unicode(fileName)
    ' Unicode (UTF-16 LE BOMあり) かどうかを判定する
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  UTF-16 LE (BOMあり) ならTrue、それ以外はFalse

    If ADODB_stream_isCharcode(fileName, "unicode") Then
        is_unicode = True
    Else
        is_unicode = False
    End If
End Function
