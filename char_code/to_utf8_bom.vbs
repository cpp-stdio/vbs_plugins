Function to_utf8_bom(fileName, CharacterCode)
    ' 指定された文字コードのファイルをUTF-8 (BOMあり) に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName      : String  変換対象ファイルのパス
    ' CharacterCode : String  変換元の文字コード
    '                         例: "unicode", "Shift_JIS", "UTF-8", "us-ascii"
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    If ADODB_stream_change_character_code(fileName, CharacterCode, "utf-8", True) Then
        to_utf8_bom = True
    Else
        to_utf8_bom = False
    End If
End Function
