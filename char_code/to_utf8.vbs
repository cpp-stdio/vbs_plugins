Function to_utf8(fileName, CharacterCode)
    ' 指定された文字コードのファイルをUTF-8 (BOMなし) に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName      : String  変換対象ファイルのパス
    ' CharacterCode : String  変換元の文字コード
    '                         例: "unicode", "Shift_JIS", "utf-8-bom", "us-ascii"
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    If ADODB_stream_change_character_code(fileName, CharacterCode, "utf-8", False) Then
        to_utf8 = True
    Else
        to_utf8 = False
    End If
End Function
