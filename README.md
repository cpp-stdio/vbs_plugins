## VBS
私自身が職場やプライベートで水平展開しても問題なしと判断したプログラムをまとめています


エクセルマクロの実行やディレクトリのタッチ、その他多くのことをたった1行で行うことができます

## Installation
サブモジュールをgitで追加するも良し、ダウンロードして利用するの良し


## Usage
利用する場合はリポジトリに "VBS "という名前を付けなければ利用できません


"VBS"がある階層でしか利用することができません





そしてプログラムを使用する際は頭に下記のコードを入力してください。
```vbscript
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```

プログラムの使い方等については各ファイルに記載してある説明文をお読みください（↓例
```vbscript
Function utf8_for_sjis(fileName, beforeLineSeparator, afterLineSeparator)
    'This program change character code
    ' UTF-8 → SJIS(SHIFT-JIS) 
    '
    'Parameters
    '----------
    'fileName : String
    '   It is the file name whose character code us want to change
    'beforeLineSeparator : String
    '   It is before change of line separator code.
    '   The line separator are as follows three patterns.
    '   [vbCrLf] : carriage return and line feed
    '   [ vbCr ] : carriage return
    '   [ vbLf ] : line feed
    'afterLineSeparator : String
    '   It is after change of line separator code.
    '   The line separator are as follows three patterns.
    '   [vbCrLf] : carriage return and line feed
    '   [ vbCr ] : carriage return
    '   [ vbLf ] : line feed
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    ...

    ...

    ...
End Function
```
