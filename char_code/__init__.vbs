'--- 文字コード判定 ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_unicode.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_utf8_bom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_utf8_nobom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_sjis.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\is_ascii.vbs").ReadAll())

'--- 文字コード変換（汎用）?変換元の文字コードを第二引数に指定 ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\to_unicode.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\to_utf8_bom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\to_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\to_sjis.vbs").ReadAll())

'--- 文字コード変換（特定パターン） ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\sjis_to_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\sjis_to_utf8_bom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\sjis_to_unicode.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\utf8_to_sjis.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\utf8_to_unicode.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\utf8_to_utf8_bom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\utf8_bom_to_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\unicode_to_utf8.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\unicode_to_utf8_bom.vbs").ReadAll())
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\unicode_to_sjis.vbs").ReadAll())

'--- private ---
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\char_code\private\__init__.vbs").ReadAll())
