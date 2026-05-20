Function to_sjis(fileName, CharacterCode)
    ' 指定された文字コードのファイルをShift_JIS に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName      : String  変換対象ファイルのパス
    ' CharacterCode : String  変換元の文字コード
    '                         例: "UTF-8", "unicode", "utf-8-bom", "us-ascii"
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    If ADODB_stream_change_character_code(fileName, CharacterCode, "Shift_JIS", False) Then
        to_sjis = True
    Else
        to_sjis = False
    End If
End Function
