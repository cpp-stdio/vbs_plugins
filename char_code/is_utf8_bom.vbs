Function is_utf8_bom(fileName)
    ' UTF-8 (BOMあり) かどうかを判定する
    ' ファイル先頭バイトが EF BB BF の場合のみTrue
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  UTF-8 BOMありならTrue、それ以外はFalse

    If ADODB_stream_isCharcode(fileName, "utf-8-bom") Then
        is_utf8_bom = True
    Else
        is_utf8_bom = False
    End If
End Function
