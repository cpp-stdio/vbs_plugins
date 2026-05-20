Function is_sjis(fileName)
    ' Shift_JIS かどうかを判定する
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  Shift_JISならTrue、それ以外はFalse

    If ADODB_stream_isCharcode(fileName, "Shift_JIS") Then
        is_sjis = True
    Else
        is_sjis = False
    End If
End Function
