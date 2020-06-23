Sub directory_move(ByVal beforeDirectoryName,ByVal afterDirectoryName)
	'Directory move
	'
	'Parameters
	'----------
	'beforeDirectoryName : String
	'	file directory before move 
	'afterDirectoryName : String
	'	file directory after move

	if beforeDirectoryName = afterDirectoryName Then Exit Sub
	Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	If objFSO.FolderExists(beforeDirectoryName) <> True Then
		WScript.Echo "Not exist, " + beforeDirectoryName
		Set objFSO = Nothing
		Exit Sub
	End if

	If objFSO.FolderExists(afterDirectoryName) <> True Then
        call creater(afterDirectoryName, objFSO)
    End If

	WScript.Echo "move " + beforeDirectoryName
	'including read-only and overwrite save
	Call objFSO.MoveFolder(beforeDirectoryName, afterDirectoryName)
	
	If Err.Number = 0 Then
		WScript.Echo "Moved to " + afterDirectoryName
	Else
		WScript.Echo "Error " + Err.Description
	End if
	Set objFSO = Nothing
End Sub
