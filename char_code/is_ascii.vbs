Function is_ascii(fileName)
    ' ASCII (全バイトが 0x7F 以下かつBOMなし) かどうかを判定する
    '
    ' Parameters
    ' ----------
    ' fileName : String  判定対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  純粋なASCIIならTrue、それ以外はFalse

    If ADODB_stream_isCharcode(fileName, "ascii") Then
        is_ascii = True
    Else
        is_ascii = False
    End If
End Function
