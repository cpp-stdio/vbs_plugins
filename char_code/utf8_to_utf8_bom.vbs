Function utf8_to_utf8_bom(fileName)
    ' UTF-8 (BOMなし) → UTF-8 (BOMあり) に変換して上書き保存する
    ' ※ すでにBOMありの場合は二重BOMにならず正常に動作する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    utf8_to_utf8_bom = to_utf8_bom(fileName, "UTF-8")
End Function
