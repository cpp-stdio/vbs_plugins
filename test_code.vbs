'When running this program, set the directory name to "VBS".

'If you want to run this program, copy it to a directory outside VBS and try again.

Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile("VBS\__init__.vbs").ReadAll())

thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))

If True = False Then
Dim directories, files, items_files
directories = get_directories(thisPath + "VBS",0)
files = get_files(thisPath + "VBS",0)
item_files = get_items(thisPath + "VBS",0)

Dim i ,str
If IsEmpty(directories) = False Then
	For i = 0 to UBound(directories)
		str = str + directories(i) + vbCrLf
	Next
	WScript.Echo str
End If

str = ""
Dim file
For Each file in files
	str = str + file + vbCrLf
Next
WScript.Echo str

str = ""
For i = 0 to UBound(item_files)
	str = str + item_files(i) + vbCrLf
Next
WScript.Echo str

End If

Call directory_contents_move(thisPath + "VBS", "C:\VBS", -1, "YES")
Call directory_contents_move("C:\VBS", thisPath + "VBS", -1, "YES")