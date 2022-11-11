## VBS
私自身が職場やプライベートで水平展開しても問題なしと判断したプログラムをまとめています\n
エクセルマクロの実行やディレクトリのタッチ、その他多くのことをたった1行で行うことができます

## Installation
サブモジュールをgitで追加するも良し、ダウンロードして利用するの良し

## Usage
利用する場合はリポジトリに "VBS "という名前を付けなければ利用できません\n
"VBA"がある階層でしか利用することはできません\n
\n
そしてプログラムを使用する際は頭に下記のコードを入力してください。
```vbscript
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```
