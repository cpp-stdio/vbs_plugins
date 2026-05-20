Function utf8_bom_to_utf8(fileName)
    ' UTF-8 (BOMあり) → UTF-8 (BOMなし) に変換して上書き保存する
    ' BOMを取り除く。すでにBOMなしの場合もそのまま正常に動作する
    '
    ' Parameters
    ' ----------
    ' fileName : String  変換対象ファイルのパス
    '
    ' Return
    ' ----------
    ' Boolean  成功(True) / 失敗(False)

    utf8_bom_to_utf8 = to_utf8(fileName, "UTF-8")
End Function
