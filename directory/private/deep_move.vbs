Function deep_move(ByVal before, ByVal after, ByRef deeper, ByRef overwrite, ByRef fso)
    deep_move = 0
    
    after = after + Right(before,len(before) - len(fso.GetParentFolderName(before)))
    Call creater(after, fso)

    Dim fsoFolder, file, folder
    Dim afterFileName
    Dim i, extension

    Set fsoFolder = fso.GetFolder(before)
    
    For Each file In fsoFolder.Files : Do
        afterFileName = after + "\" + file.name

        If fso.FileExists(afterFileName) = True Then
            If StrComp(overwrite, "YES", 1) = 0 Then
                Call fso.DeleteFile(afterFileName, True)
            ElseIf StrComp(overwrite, "ANOTHER", 1) = 0 Then
                i = 1
                Do
                    extension = fso.GetExtensionName(afterFileName)
                    If len(extension) > 0 Then
                        afterFileName = after + "\" + left(file.name, len(file.name) - len(extension) - 1)
                        afterFileName = afterFileName + "(" + Cstr(i) + ")." + extension
                    Else
                        afterFileName = after + "\" + file.name + "(" + Cstr(i) + ")"
                    End If

                    If fso.FileExists(afterFileName) = False Then Exit Do
                    i = i + 1
                Loop
            ElseIf StrComp(overwrite, "NO", 1) = 0 Then
                Exit Do 'mimic continue
            End If
        End If
        Call fso.MoveFile(before + "\" + file.name, afterFileName)
    Loop Until 1 : Next

    If deeper <= -1 Then
        deeper = -1
    End if

    Dim remain: remain = 0
    If deeper <> 0 Then
        For Each folder In fsoFolder.subfolders
            remain = remain + deep_move(before + "\" + folder.name, after, deeper - 1, overwrite, fso)
        Next
    End If
    
    Set fsoFolder = fso.GetFolder(before)
    remain = remain + fsoFolder.Files.Count
    remain = remain + fsoFolder.subfolders.Count
    
    If remain <= 0 Then
        call fso.DeleteFolder(before, True)
    End If

    deep_move = remain
    Set fsoFolder = Nothing
    Set file = Nothing
    Set folder = Nothing
End Function
