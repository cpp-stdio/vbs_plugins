Sub directory_contents_copy(ByVal beforeDirectoryName,ByVal afterDirectoryName,ByVal deeper)
	'It copy directory contents
	'
	'Parameters
	'----------
	'beforeDirectoryName : String
	'	directory path before copy 
	'afterDirectoryName : String
	'	directory path after copy
	'deeper : int
	'	[Negative number] It copy all directory contents
	'	[       0       ] It copy a directory contents
	'	[positive number] It copies directory contents deep hierarchy for the according to number

	If beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If Right(beforeDirectoryName,1) = "\" Then
		beforeDirectoryName = left(beforeDirectoryName, len(beforeDirectoryName) - 1)
	End If

	If Right(afterDirectoryName,1) = "\" Then
		afterDirectoryName = left(afterDirectoryName, len(afterDirectoryName) - 1)
	End If

	On Error Resume Next
	If objFso.FolderExists(beforeDirectoryName) <> True Then
		WScript.Echo "Not exist, " + beforeDirectoryName
		Set objFSO = Nothing
		Exit Sub
	End If

	If objFSO.FolderExists(afterDirectoryName) <> True Then
        call creater(afterDirectoryName, objFso)
    End If

	Call deep_copy(beforeDirectoryName, afterDirectoryName, deeper, objFSO)

	If Err.Number = 0 Then
		WScript.Echo "Copied to " + afterDirectoryName
	Else
		WScript.Echo "Error " + Err.Description
	End if

	Set objFSO = Nothing
End Sub
