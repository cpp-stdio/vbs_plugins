Function delete_file(ByVal fileName)
    'Delete fileName
    '
    'Parameters
    '----------
    'directoryName : String
    '   file name to be delete
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    delete_file = False

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    
    If objFso.FileExists(fileName) = True Then
        'Delete directory(including read-only and subdirectory)
        call objFso.DeleteFile(fileName, True)

        If Err.Number <> 0 Then
            WScript.Echo "The file being edited exists."
        Else
            WScript.Echo "completed, Deletion of " + fileName
            delete_file = True
        End if
    Else
        WScript.Echo "Not exist, " + fileName
    End If

    Set objFso = Nothing
End Function

