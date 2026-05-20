Function deep_creater(ByVal path, ByRef fso)
    Dim strParent: strParent = fso.GetParentFolderName(path)

    If fso.FolderExists(strParent) = True Then
        If fso.FolderExists(path) <> True Then
            Call fso.CreateFolder(path)
        End If
    Else
        Call deep_creater(strParent, fso)
        Call fso.CreateFolder(path)
    End If

    Set strParent = Nothing
End Function
