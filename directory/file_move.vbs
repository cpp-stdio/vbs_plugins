Sub file_move(ByVal beforeFileName,ByVal afterFileName)
    'File move
    '
    'Parameters
    '----------
    'beforeFileName : String
    '   file name before move 
    'afterFileName : String
    '   file name after move

    if beforeFileName = afterFileName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    If objFso.FileExists(beforeFileName) Then

        WScript.Echo "move " + beforeFileName
        'including read-only and overwrite save
        Call objFSO.MoveFile(beforeFileName, afterFileName)
        WScript.Echo "Moved to " + afterFileName
    Else
        WScript.Echo "Not exist, " + beforeFileName
    End If
    Set objFso = Nothing
End Sub
