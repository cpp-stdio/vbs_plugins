Function deep_items(ByVal directoryName, ByRef deeper, ByRef element, ByRef fso)

    Dim items()
    Dim length: length = 0
    deep_items = items

    Dim fsoFolder, file, folder
    Set fsoFolder = fso.GetFolder(directoryName)

    If element = "FILE" Or element = "ALL" Then
        For Each file In fsoFolder.Files
            ReDim Preserve items(length)
            items(length) = directoryName + "\" + file.name
            length = length + 1
        Next
    End If

    Set file = Nothing
    If deeper <= -1 Then
        deeper = -1
    End if
    
    Dim dict ,dicts
    For Each folder In fsoFolder.subfolders
        If element = "FOLDER" Or element = "ALL" Then
            ReDim Preserve items(length)
            items(length) = directoryName + "\" + folder.name
            length = length + 1
        End If

        If deeper <> 0 Then
            dicts = deep_items(directoryName + "\" + folder.name, deeper, element, fso)
            For Each dict In dicts
                ReDim Preserve items(length)
                items(length) = dict
                length = length + 1
            Next
        End If
    Next

    Set fsoFolder = Nothing
    Set folder = Nothing
    deep_items = items
End Function
