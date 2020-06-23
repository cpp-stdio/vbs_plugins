Sub directory_contents_move(ByVal beforeDirectoryName,ByVal afterDirectoryName,ByVal deeper,ByVal overwrite)
	'It move directory contents
	'
	'Parameters
	'----------
	'beforeDirectoryName : String
	'	It directory path before move
	'afterDirectoryName : String
	'	It directory path after move
	'deeper : int
	'	[Negative number] It move all directory contents
	'	[       0       ] It move a directory contents
	'	[positive number] It moves directory contents deep hierarchy for the according to number
	'overwrite : String
	'	If data with the same name exists for afterDirectoryName, do you want to overwrite it?
	'	"YES" 	  : Yes, do it
	'	"DELETE"  : Yes, do it. And if there is afterDirectoryName, delete all of the directory and execute.
	'	"ANOTHER" : Yes, do it. But please another name
	'	"NO"      : No, don't it. Please, ignore

	If beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If Right(beforeDirectoryName,1) = "\" Then
		beforeDirectoryName = left(beforeDirectoryName, len(beforeDirectoryName) - 1)
	End If

	If Right(afterDirectoryName,1) = "\" Then
		afterDirectoryName = left(afterDirectoryName, len(afterDirectoryName) - 1)
	End If

	On Error Resume Next
	If objFSO.FolderExists(beforeDirectoryName) <> True Then
		WScript.Echo "Not exist, " + beforeDirectoryName
		Set objFSO = Nothing
		Exit Sub
	End If

	If StrComp(overwrite, "DELETE", 1) = 0 Then
		Dim delete_path
		delete_path = afterDirectoryName + Right(beforeDirectoryName,len(beforeDirectoryName) - len(objFSO.GetParentFolderName(beforeDirectoryName)))
		
		If objFSO.FolderExists(delete_path) = True Then 
			Call delete_directory(delete_path)
		End If
		overwrite = "YES"
	End If

	Call deep_move(beforeDirectoryName, afterDirectoryName, deeper, overwrite, objFSO)

	If Err.Number = 0 Then
		WScript.Echo "Moved to " + afterDirectoryName
	Else
		WScript.Echo "Error " + Err.Description
	End if

	Set objFSO = Nothing
End Sub
