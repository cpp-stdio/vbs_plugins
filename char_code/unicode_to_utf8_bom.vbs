Function unicode_to_utf8_bom(fileName)
    ' Unicode (UTF-16 LE) → UTF-8 (BOMあり) に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    unicode_to_utf8_bom = to_utf8_bom(fileName, "unicode")
End Function
