Sub directory_copy(ByVal beforeDirectoryName,ByVal afterDirectoryName)
	'Directory copy
	'
	'Parameters
	'----------
	'beforeDirectoryName : String
	'	file directory before copy 
	'afterDirectoryName : String
	'	file directory after copy

	if beforeDirectoryName = afterDirectoryName Then Exit Sub
	Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	If objFso.FolderExists(beforeDirectoryName) <> True Then
		WScript.Echo "Not exist, " + beforeDirectoryName
		Set objFso = Nothing
		Exit Sub
	End if

	If objFSO.FolderExists(afterDirectoryName) <> True Then
        call creater(afterDirectoryName, objFso)
    End If

	WScript.Echo "copy " + beforeDirectoryName
	'including read-only and overwrite save
	Call objFSO.CopyFolder(beforeDirectoryName, afterDirectoryName, True)
	
	If Err.Number = 0 Then
		WScript.Echo "Copied to " + afterDirectoryName
	Else
		WScript.Echo "Error " + Err.Description
	End if
	Set objFso = Nothing
End Sub
