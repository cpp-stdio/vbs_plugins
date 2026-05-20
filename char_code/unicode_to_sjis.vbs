Function unicode_to_sjis(fileName)
    ' Unicode (UTF-16 LE) → Shift_JIS に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    unicode_to_sjis = to_sjis(fileName, "unicode")
End Function
