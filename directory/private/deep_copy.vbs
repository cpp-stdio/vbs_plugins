Function deep_copy(ByVal before, ByVal after, ByVal deeper, ByRef fso)
	
    If fso.FolderExists(after) <> True Then
	    fso.CreateFolder(after)
    End If

	Dim fsoFolder, file, folder
	Set fsoFolder = fso.GetFolder(before)
	For Each file In fsoFolder.Files
		Call fso.CopyFile(before + "\" + file.name, after + "\" + file.name, True)
	Next

	if deeper = 0 Then Exit Function
	if deeper <= -1 Then
		deeper = -1
	End if

	For Each folder In fsoFolder.subfolders
		Call deep_copy(before + "\" + folder.name, after + "\" + folder.name, deeper - 1, fso)
	Next 

	Set fsoFolder = Nothing
	Set file = Nothing
	Set folder = Nothing
End Function
