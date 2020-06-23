'public
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\sjis_for_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\utf8_for_sjis.vbs").ReadAll())

'private
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\private\__init__.vbs").ReadAll())