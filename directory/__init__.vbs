' Loads all directory operation functions.
' ディレクトリ操作に関するすべての関数を読み込む。

' --- Public / 公開関数 ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\create_directory.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\delete_directory.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\delete_file.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_contents_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_contents_move.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_move.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\file_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\file_move.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_directories.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_files.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_items.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\up_folder.vbs").ReadAll())

' --- Private / 内部実装 ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\__init__.vbs").ReadAll())
