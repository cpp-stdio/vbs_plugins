Function delete_directory(ByVal directoryName)
	'Delete all its files and folders, including directory
	'
	'Parameters
	'----------
	'directoryName : String
	'	directory name to be delete
	'
	'Return
	'----------
	'boolen
	'	success(True) , failure(False)
	delete_directory = False

	Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	
	If objFso.FolderExists(directoryName) = True Then
		'Delete directory(including read-only and subdirectory)
		call objFso.DeleteFolder(directoryName, True)

		If Err.Number <> 0 Then
			WScript.Echo "The directory being edited exists."
		Else
			WScript.Echo "completed, Deletion of " + directoryName
			delete_directory = True
		End if
	Else
		WScript.Echo "Not exist, " + directoryName
	End If

	Set objFso = Nothing
End Function
