Function is_utf8_nobom(fileName)
    ' UTF-8 (BOMなし) かどうかを判定する
    ' BOMなし かつ 有効なUTF-8バイト列の場合のみTrue
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  UTF-8 BOMなしならTrue、BOMありまたは非UTF-8はFalse

    If ADODB_stream_isCharcode(fileName, "utf-8-nobom") Then
        is_utf8_nobom = True
    Else
        is_utf8_nobom = False
    End If
End Function
