Function utf8_to_unicode(fileName)
    ' UTF-8 (BOMあり・なし問わず) → Unicode (UTF-16 LE BOMあり) に変換して上書き保存する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    utf8_to_unicode = to_unicode(fileName, "UTF-8")
End Function
