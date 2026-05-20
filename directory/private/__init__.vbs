' Private helper functions for directory operations.
' These are for internal use only and are not intended to be called directly from outside.
' ディレクトリ操作の内部実装。外部から直接呼び出すことは想定していない。
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\deep_creater.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\deep_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\deep_items.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\deep_move.vbs").ReadAll())
