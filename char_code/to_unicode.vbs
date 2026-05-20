Function to_unicode(fileName, CharacterCode)
    ' 指定された文字コードのファイルをUnicode (UTF-16 LE BOMあり) に変換して上書き保存する
    ' ※ Unicode (UTF-16 LE) のBOMはADODB.Streamが自動で付与する
    '
    ' Parameters
    ' ----------
    ' fileName      : String  変換対象ファイルのパス
    ' CharacterCode : String  変換元の文字コード
    '                         例: "UTF-8", "Shift_JIS", "utf-8-bom", "us-ascii"
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    If ADODB_stream_change_character_code(fileName, CharacterCode, "unicode", False) Then
        to_unicode = True
    Else
        to_unicode = False
    End If
End Function
