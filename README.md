## VBS
私自身が職場やプライベートで水平展開しても問題なしと判断したプログラムをまとめています。
エクセルマクロの実行やディレクトリのタッチ、その他多くのことをたった1行で行うことができます

## Installation
サブモジュールをgitで追加するも良し、ダウンロードして利用するの良し

## Usage
利用する場合はリポジトリに "VBS "という名前を付けなければ利用できません
本来のVBSにC++のようなインクルードという仕組みはありません。しかし抜け道がありファイルの階層を理解していれば実は可能です
```vbscript
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```
