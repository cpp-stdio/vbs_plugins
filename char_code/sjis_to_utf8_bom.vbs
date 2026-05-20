Function sjis_to_utf8_bom(fileName)
    ' Shift_JIS → UTF-8 (BOMあり) に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    sjis_to_utf8_bom = to_utf8_bom(fileName, "Shift_JIS")
End Function
