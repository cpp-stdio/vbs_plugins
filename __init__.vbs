' Main entry point for the VBScript utility library.
' Loads all public modules. Execute this file at the start of your script to make all functions available.
' VBScriptユーティリティライブラリのメインエントリーポイント。
' スクリプトの先頭でこのファイルを Execute することで、すべての関数が利用可能になる。

' --- Public modules / 公開モジュール ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\__init__.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\excel\__init__.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\__init__.vbs").ReadAll())

' --- Private ---
