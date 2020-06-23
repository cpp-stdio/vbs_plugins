Sub file_copy(ByVal beforeFileName,ByVal afterFileName)
	'File copy
	'
	'Parameters
	'----------
	'beforeFileName : String
	'	file name before copy 
	'afterFileName : String
	'	file name after copy

	if beforeFileName = afterFileName Then Exit Sub
	Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If objFso.FileExists(beforeFileName) Then

		WScript.Echo "copy " + beforeFileName
		'including read-only and overwrite save
		Call objFSO.CopyFile(beforeFileName, afterFileName, True)
		WScript.Echo "Copied to " + afterFileName
	Else
		WScript.Echo "Not exist, " + beforeFileName
	End If
	Set objFso = Nothing
End Sub
