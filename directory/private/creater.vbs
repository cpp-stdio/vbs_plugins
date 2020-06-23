Function creater(ByVal path, ByRef fso)
    
    Dim strParent: strParent = fso.GetParentFolderName(path)

    If fso.FolderExists(strParent) = True Then
        If fso.FolderExists(path) <> True Then
            call fso.CreateFolder(path)
        End If
    Else
        Call creater(strParent, fso)
        fso.CreateFolder(path)
    End If

    Set strParent = Nothing
End Function
