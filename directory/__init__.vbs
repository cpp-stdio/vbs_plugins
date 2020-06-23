'public
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\create_directory.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\delete_directory.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\delete_file.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_contents_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_contents_move.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\directory_move.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\file_copy.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_directories.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_files.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\get_items.vbs").ReadAll())

'private
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\directory\private\__init__.vbs").ReadAll())
